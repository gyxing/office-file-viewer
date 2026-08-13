import type { OfficeArchiveReader } from '../../shared/ooxml/OfficeArchiveReader';
import { OFFICE_LARGE_FILE_THRESHOLDS } from '../performance/officePerformanceThresholds';
import { WordPerformanceStatsCollector } from '../word/collectWordPerformanceStats';
import { createProgressiveWordOutlineProvider } from '../word/createMemoryWordOutlineProvider';
import { createWordPerformanceProfile } from '../word/performance';
import type { WordOutlineItem } from '../word/types';
import type { WordOutlineProvider } from '../word/WordOutlineProvider';
import type {
  WordPageMeta,
  WordPageSource,
  WordPageSourceSnapshot,
} from '../word/WordPageSource';
import { WordPageStore } from '../word/WordPageStore';
import type { WordPreviewSource } from '../word/WordPreviewSource';
import {
  collectDocxSearchBlocks,
  WordSearchProvider,
} from '../word/WordSearchProvider';
import {
  createDocxMeasurementBatches,
  paginateMeasuredDocxPage,
  type DocxMeasuredBlock,
  type DocxMeasurementBatch,
} from './docxPagination';
import type { DocxDocument, DocxPage, DocxPageContent } from './types';

/** DOCX 页面数据源的预览摘要。 */
export type DocxWordPreviewSummary = Omit<DocxDocument, 'blocks' | 'pages'>;

/** 创建 DOCX 页面数据源时使用的选项。 */
type DocxWordPageSourceOptions = {
  /** 当前解析或预览会话的标识。 */
  sessionId: string;
  /** 用于按需读取源数据的读取器。 */
  reader: OfficeArchiveReader;
  /** 用于取消当前异步操作的信号。 */
  signal?: AbortSignal;
  /** 报告不会阻断预览的解析或缓存警告。 */
  onWarning?(error: unknown): void;
};

/** 等待完成解析和分页的 DOCX 源页面。 */
type PendingSourcePage = {
  /** 分页处理前所属的源页面。 */
  sourcePage: DocxPageContent;
  /** 按处理顺序排列的内容批次。 */
  batches: DocxMeasurementBatch[];
  /** 当前源页面已经完成的内容块测量结果。 */
  measurements: DocxMeasuredBlock[];
};

/** DOCX Viewer 可同时消费主线程 Source 与 Worker 代理的稳定接口。 */
export interface DocxPagePreviewSource
  extends WordPageSource<DocxPageContent>,
    WordPreviewSource<DocxPageContent> {
  /** 返回当前文档摘要。 */
  getSummary(): DocxWordPreviewSummary;
  /** 返回当前已经发现的扁平大纲条目。 */
  getOutlineItems(): WordOutlineItem[];
  /** 返回当前等待浏览器布局测量的批次。 */
  getMeasurementBatch(): DocxMeasurementBatch | undefined;
  /** 提交主线程测得的块尺寸并推进分页。 */
  commitMeasurement(
    batch: DocxMeasurementBatch,
    measurements: readonly DocxMeasuredBlock[],
    durationMs: number,
  ): Promise<void>;
  /** 标记当前测量批次失败，供界面触发重试。 */
  failMeasurement(batch: DocxMeasurementBatch, error: unknown): void;
  /** 当前 Source 是否已经具备可渲染摘要。 */
  hasRenderableContent(): boolean;
}

function estimatePageSize(page: DocxPageContent) {
  return page.blocks.reduce((total, block) => {
    if (block.type === 'paragraph') return total + block.text.length * 2 + 256;
    if (block.type === 'table') return total + block.rows.length * 512;
    return total + 2048;
  }, 512);
}

/** 流式 DOCX 的页面、测量队列、大纲和懒归档资源的唯一所有者。 */
export class DocxWordPageSource implements DocxPagePreviewSource {
  readonly pages: WordPageSource<DocxPageContent> = this;
  readonly outline: WordOutlineProvider;
  private readonly writableOutline = createProgressiveWordOutlineProvider();
  private readonly reader: OfficeArchiveReader;
  private readonly signal?: AbortSignal;
  private readonly pageStore: WordPageStore<DocxPageContent>;
  private readonly stats = new WordPerformanceStatsCollector();
  private readonly listeners = new Set<() => void>();
  private readonly outlineItemsById = new Map<string, WordOutlineItem>();
  private readonly blockPageIndex = new Map<string, number>();
  readonly searchProvider = new WordSearchProvider({
    resolvePageIndex: (blockId) => this.blockPageIndex.get(blockId),
  });
  private readonly batchById = new Map<string, DocxMeasurementBatch>();
  private readonly pendingPages = new Map<string, PendingSourcePage>();
  private readonly measurementQueue: DocxMeasurementBatch[] = [];
  private snapshot: WordPageSourceSnapshot = { revision: 0, pages: [] };
  private summary?: DocxWordPreviewSummary;
  private nextBatchRevision = 1;
  private parsingCompleted = false;
  private completed = false;
  private disposed = false;
  private failedBatchId?: string;
  private disposePromise?: Promise<void>;

  constructor(options: DocxWordPageSourceOptions) {
    this.reader = options.reader;
    this.signal = options.signal;
    this.outline = this.writableOutline;
    this.pageStore = new WordPageStore({
      sessionId: options.sessionId,
      estimateSize: estimatePageSize,
      onWarning: options.onWarning,
    });
  }

  setMetadata(metadata: {
    page: DocxPage;
    preserveSectionPagination: boolean;
    characterSpacingControl?: DocxDocument['characterSpacingControl'];
  }) {
    this.throwIfUnavailable();
    this.summary = {
      title: '',
      page: metadata.page,
      images: [],
      outline: [],
      bookmarks: {},
      preserveSectionPagination: metadata.preserveSectionPagination,
      characterSpacingControl: metadata.characterSpacingControl,
    };
    this.emitChange();
  }

  async addSourcePage(page: DocxPageContent) {
    this.throwIfUnavailable();
    this.collectPageMetadata(page);
    this.searchProvider.append(collectDocxSearchBlocks(page.blocks));
    // 每个 sourcePage 仍独立分页，因此测量不会跨越显式节边界。
    const batches = createDocxMeasurementBatches(page, this.nextBatchRevision);
    this.nextBatchRevision += batches.length;
    if (!batches.length) {
      await this.appendReadyPages([page]);
      return;
    }
    this.pendingPages.set(page.id, {
      sourcePage: page,
      batches,
      measurements: [],
    });
    for (const batch of batches) {
      await this.waitForQueueCapacity();
      this.batchById.set(batch.id, batch);
      this.measurementQueue.push(batch);
      this.snapshot = {
        revision: this.snapshot.revision + 1,
        pages: [
          ...this.snapshot.pages,
          this.createBatchMeta(batch, 'estimated'),
        ],
      };
      this.emitChange();
    }
  }

  finishParsing(result: {
    title: string;
    images: DocxDocument['images'];
    bookmarks: NonNullable<DocxDocument['bookmarks']>;
  }) {
    this.throwIfUnavailable();
    if (!this.summary) throw new Error('DOCX Source 尚未收到元数据');
    this.summary = {
      ...this.summary,
      title: result.title,
      images: result.images,
      bookmarks: result.bookmarks,
    };
    this.parsingCompleted = true;
    this.searchProvider.complete();
    this.tryComplete();
    this.emitChange();
  }

  getMeasurementBatch() {
    return this.failedBatchId ? undefined : this.measurementQueue[0];
  }

  async commitMeasurement(
    batch: DocxMeasurementBatch,
    measurements: readonly DocxMeasuredBlock[],
    durationMs: number,
  ) {
    if (this.disposed || this.measurementQueue[0]?.id !== batch.id) return;
    const pending = this.pendingPages.get(batch.sourcePage.id);
    if (!pending) return;
    const previousMeasurementCount = pending.measurements.length;
    pending.measurements.push(...measurements);
    if (durationMs >= OFFICE_LARGE_FILE_THRESHOLDS.slowTaskMilliseconds) {
      this.stats.reportSlowPagination();
    }
    if (batch.endOfSourcePage) {
      try {
        const pages = paginateMeasuredDocxPage(
          pending.sourcePage,
          pending.measurements,
        );
        await this.replaceBatchesWithPages(pending.batches, pages);
      } catch (error) {
        // 末批分页或持久化失败时恢复队列状态，避免同一测量结果在重试时重复追加。
        pending.measurements.length = previousMeasurementCount;
        throw error;
      }
      this.pendingPages.delete(batch.sourcePage.id);
      pending.batches.forEach((item) => this.batchById.delete(item.id));
    }
    this.measurementQueue.shift();
    this.tryComplete();
    this.emitChange();
  }

  failMeasurement(batch: DocxMeasurementBatch, error: unknown) {
    if (this.disposed || this.measurementQueue[0]?.id !== batch.id) return;
    this.measurementQueue.shift();
    this.failedBatchId = batch.id;
    const message =
      error instanceof Error ? error.message : 'DOCX 页面测量失败';
    this.snapshot = {
      revision: this.snapshot.revision + 1,
      pages: this.snapshot.pages.map((meta) =>
        meta.id === batch.id
          ? {
              ...meta,
              revision: meta.revision + 1,
              status: 'error',
              errorMessage: message,
            }
          : meta,
      ),
    };
    this.emitChange();
  }

  hasRenderableContent() {
    return Boolean(this.summary);
  }

  getSummary() {
    if (!this.summary) throw new Error('DOCX Source 尚无文档摘要');
    return this.summary;
  }

  getOutlineItems() {
    return [...this.outlineItemsById.values()];
  }

  getPerformanceProfile() {
    const profile = createWordPerformanceProfile(this.stats.getSnapshot());
    return { ...profile, pageMode: 'windowed' as const };
  }

  getSnapshot = () => this.snapshot;

  subscribe = (listener: () => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  async getPage(index: number, signal?: AbortSignal) {
    const meta = this.snapshot.pages[index];
    if (!meta || meta.status !== 'ready') return undefined;
    return this.pageStore.get(meta, signal);
  }

  async ensureRange(start: number, end: number, signal?: AbortSignal) {
    if (end <= start) return;
    await this.waitUntil(
      () => this.snapshot.pages.length >= end || this.completed,
      signal,
    );
  }

  async prioritizeBlock(blockId: string, signal?: AbortSignal) {
    await this.waitUntil(
      () => this.blockPageIndex.has(blockId) || this.completed,
      signal,
    );
    return this.blockPageIndex.get(blockId) ?? -1;
  }

  retainRange(start: number, end: number) {
    return this.pageStore.retain(this.snapshot.pages.slice(start, end));
  }

  retry(index: number) {
    const meta = this.snapshot.pages[index];
    const batch = meta ? this.batchById.get(meta.id) : undefined;
    if (!batch || meta.status !== 'error' || this.failedBatchId !== batch.id) {
      return;
    }
    this.failedBatchId = undefined;
    this.measurementQueue.unshift(batch);
    this.snapshot = {
      revision: this.snapshot.revision + 1,
      pages: this.snapshot.pages.map((item) =>
        item.id === batch.id
          ? {
              ...item,
              revision: item.revision + 1,
              status: 'estimated',
              errorMessage: undefined,
            }
          : item,
      ),
    };
    this.emitChange();
  }

  async waitForCompletion(signal?: AbortSignal) {
    await this.waitUntil(() => this.completed, signal);
  }

  dispose() {
    if (this.disposePromise) return this.disposePromise;
    this.disposed = true;
    this.completed = true;
    this.searchProvider.dispose();
    this.writableOutline.complete();
    this.emitChange();
    this.listeners.clear();
    this.measurementQueue.length = 0;
    this.batchById.clear();
    this.pendingPages.clear();
    this.blockPageIndex.clear();
    this.outlineItemsById.clear();
    this.disposePromise = Promise.all([
      this.pageStore.dispose(),
      this.reader.close(),
    ]).then(() => undefined);
    return this.disposePromise;
  }

  private collectPageMetadata(page: DocxPageContent) {
    this.stats.addDocxBlocks(page.blocks);
    page.blocks.forEach((block) => {
      if (
        block.type !== 'paragraph' ||
        block.outlineLevel === undefined ||
        !block.text
      ) {
        return;
      }
      const item = {
        id: `outline-${block.id}`,
        text: block.text,
        level: block.outlineLevel,
        targetBlockId: block.id,
      };
      if (this.outlineItemsById.has(item.id)) return;
      this.outlineItemsById.set(item.id, item);
      this.writableOutline.append([item]);
      this.stats.setOutlineCount(this.outlineItemsById.size);
    });
  }

  private createBatchMeta(
    batch: DocxMeasurementBatch,
    status: WordPageMeta['status'],
  ): WordPageMeta {
    return {
      id: batch.id,
      index: this.snapshot.pages.length,
      revision: batch.revision,
      status,
      estimatedContentHeight: batch.sourcePage.page.minHeight,
      sourceBlockIds: batch.blocks.flatMap((block) => [
        block.id,
        ...(block.sourceBlockId ? [block.sourceBlockId] : []),
      ]),
    };
  }

  private async appendReadyPages(pages: readonly DocxPageContent[]) {
    const metas: WordPageMeta[] = [];
    for (const page of pages) {
      const index = this.snapshot.pages.length + metas.length;
      const meta = this.createReadyMeta(page, index);
      await this.pageStore.put(meta, page);
      metas.push(meta);
    }
    this.snapshot = {
      revision: this.snapshot.revision + 1,
      pages: [...this.snapshot.pages, ...metas],
    };
    metas.forEach((meta) => {
      meta.sourceBlockIds.forEach((id) =>
        this.blockPageIndex.set(id, meta.index),
      );
    });
    this.emitChange();
  }

  private async replaceBatchesWithPages(
    batches: readonly DocxMeasurementBatch[],
    pages: readonly DocxPageContent[],
  ) {
    const batchIds = new Set(batches.map((batch) => batch.id));
    const firstIndex = this.snapshot.pages.findIndex((meta) =>
      batchIds.has(meta.id),
    );
    const readyMetas: WordPageMeta[] = [];
    for (const page of pages) {
      const meta = this.createReadyMeta(
        page,
        Math.max(0, firstIndex) + readyMetas.length,
      );
      await this.pageStore.put(meta, page);
      readyMetas.push(meta);
    }
    const remaining = this.snapshot.pages.filter(
      (meta) => !batchIds.has(meta.id),
    );
    const insertionIndex = Math.max(0, firstIndex);
    const nextPages = [
      ...remaining.slice(0, insertionIndex),
      ...readyMetas,
      ...remaining.slice(insertionIndex),
    ].map((meta, index) => ({ ...meta, index }));
    this.snapshot = {
      revision: this.snapshot.revision + 1,
      pages: nextPages,
    };
    this.rebuildBlockIndex();
  }

  private createReadyMeta(page: DocxPageContent, index: number): WordPageMeta {
    return {
      id: page.id,
      index,
      revision: 1,
      status: 'ready',
      estimatedContentHeight: page.page.minHeight,
      sourceBlockIds: page.blocks.flatMap((block) => [
        block.id,
        ...(block.sourceBlockId ? [block.sourceBlockId] : []),
      ]),
    };
  }

  private rebuildBlockIndex() {
    this.blockPageIndex.clear();
    this.snapshot.pages.forEach((meta) => {
      if (meta.status !== 'ready') return;
      meta.sourceBlockIds.forEach((id) =>
        this.blockPageIndex.set(id, meta.index),
      );
    });
  }

  private tryComplete() {
    if (
      this.completed ||
      !this.parsingCompleted ||
      this.measurementQueue.length ||
      this.failedBatchId ||
      this.pendingPages.size
    ) {
      return;
    }
    this.completed = true;
    this.writableOutline.complete();
    this.stats.setEstimatedPageCount(this.snapshot.pages.length);
    this.snapshot = {
      ...this.snapshot,
      revision: this.snapshot.revision + 1,
      pageCount: this.snapshot.pages.length,
    };
  }

  private async waitForQueueCapacity() {
    await this.waitUntil(
      () =>
        this.measurementQueue.length + (this.failedBatchId ? 1 : 0) < 6 ||
        this.disposed,
      this.signal,
    );
  }

  private async waitUntil(predicate: () => boolean, signal?: AbortSignal) {
    this.throwIfAborted(signal);
    if (predicate()) return;
    await new Promise<void>((resolve, reject) => {
      let unsubscribe: () => void = () => undefined;
      let abort: () => void = () => undefined;
      const cleanup = () => {
        unsubscribe();
        signal?.removeEventListener('abort', abort);
        this.signal?.removeEventListener('abort', abort);
      };
      const resolveWhenReady = () => {
        if (!predicate()) return;
        cleanup();
        resolve();
      };
      abort = () => {
        cleanup();
        const error = new Error('DOCX Source 操作已取消');
        error.name = 'AbortError';
        reject(error);
      };
      unsubscribe = this.subscribe(resolveWhenReady);
      signal?.addEventListener('abort', abort, { once: true });
      this.signal?.addEventListener('abort', abort, { once: true });
      if (signal?.aborted || this.signal?.aborted) abort();
    });
  }

  private throwIfAborted(signal?: AbortSignal) {
    if (signal?.aborted || this.signal?.aborted) {
      const error = new Error('DOCX Source 操作已取消');
      error.name = 'AbortError';
      throw error;
    }
  }

  private throwIfUnavailable() {
    this.throwIfAborted();
    if (this.disposed) throw new Error('DOCX Source 已释放');
  }

  private emitChange() {
    this.listeners.forEach((listener) => listener());
  }
}
