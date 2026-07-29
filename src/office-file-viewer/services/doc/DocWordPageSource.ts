import {
  DocPaginationState,
  type PaginatedDocPage,
} from '../../formats/doc/docRenderUtils';
import {
  createTimeBudget,
  throwIfAborted,
  yieldToMainThread,
} from '../../formats/word-pages/scheduler';
import {
  resolveDocBlockResources,
  resolveDocMetadataResources,
} from '../parsing/assembly/DocumentAssembler';
import { ResourceRegistry } from '../parsing/assembly/ResourceRegistry';
import type {
  PortableDocMetadata,
  PortableResource,
} from '../parsing/protocol/messages';
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
import { estimateDocBlockBytes } from './chunkDocBlocks';
import type { DocBlock, DocDocument } from './types';

export type DocWordPreviewSummary = Omit<
  DocDocument,
  'blocks' | 'paragraphs' | 'resources'
>;

type DocWordPageSourceOptions = {
  sessionId: string;
  signal?: AbortSignal;
  onWarning?(error: unknown): void;
};

/** 直接消费 DOC 解析分块并把完成页写入共享页面 Store。 */
export class DocWordPageSource
  implements
    WordPageSource<PaginatedDocPage>,
    WordPreviewSource<PaginatedDocPage>
{
  readonly pages: WordPageSource<PaginatedDocPage> = this;
  readonly outline: WordOutlineProvider;
  private readonly writableOutline = createProgressiveWordOutlineProvider();
  private readonly resources = new ResourceRegistry();
  private readonly pageStore: WordPageStore<PaginatedDocPage>;
  private readonly stats = new WordPerformanceStatsCollector();
  private readonly listeners = new Set<() => void>();
  private readonly blockPageIndex = new Map<string, number>();
  private readonly outlineItemsById = new Map<string, WordOutlineItem>();
  private readonly signal?: AbortSignal;
  private pagination?: DocPaginationState;
  private summary?: DocWordPreviewSummary;
  private snapshot: WordPageSourceSnapshot = {
    revision: 0,
    pages: [],
  };
  private expectedBlockIndex = 0;
  private completed = false;
  private disposePromise?: Promise<void>;

  constructor(options: DocWordPageSourceOptions) {
    this.signal = options.signal;
    this.outline = this.writableOutline;
    this.pageStore = new WordPageStore({
      sessionId: options.sessionId,
      estimateSize: (page) =>
        page.blocks.reduce(
          (sum, block) => sum + estimateDocBlockBytes(block),
          0,
        ),
      onWarning: options.onWarning,
    });
  }

  addResource(resource: PortableResource) {
    return this.resources.register(resource);
  }

  setMetadata(metadata: PortableDocMetadata) {
    throwIfAborted(this.signal);
    const resolved = resolveDocMetadataResources(metadata, this.resources);
    this.summary = resolved;
    this.stats.setOutlineCount(resolved.outline?.length ?? 0);
    resolved.images.forEach((image) => this.stats.addImage(image.id));
    if (resolved.headerImage) this.stats.addImage(resolved.headerImage.id);
    this.appendOutline(resolved.outline ?? []);
    if (!this.pagination) {
      const contentWidth =
        resolved.page.width -
        resolved.page.marginLeft -
        resolved.page.marginRight;
      this.pagination = new DocPaginationState(resolved.page, contentWidth);
    }
    this.emitChange();
  }

  async addBlocks(startIndex: number, blocks: DocBlock[]) {
    throwIfAborted(this.signal);
    if (!this.pagination || !this.summary) {
      throw new Error('DOC PageSource 尚未收到文档元数据');
    }
    if (startIndex !== this.expectedBlockIndex) {
      throw new Error(
        `DOC 正文分块顺序无效：期望 ${this.expectedBlockIndex}，收到 ${startIndex}`,
      );
    }
    const budget = createTimeBudget();
    for (const block of blocks) {
      resolveDocBlockResources(block, this.resources);
      this.stats.addDocBlocks([block]);
      if (
        block.type === 'paragraph' &&
        block.outlineLevel !== undefined &&
        block.text
      ) {
        this.appendOutline([
          {
            id: `outline-${block.id}`,
            text: block.text,
            level: block.outlineLevel,
            targetBlockId: block.id,
          },
        ]);
      }
      for (const page of this.pagination.append([block])) {
        await this.publishPage(page);
      }
      this.expectedBlockIndex += 1;
      if (budget.shouldYield()) {
        await yieldToMainThread(this.signal);
        budget.reset();
      }
    }
  }

  async complete(warnings?: string[]) {
    if (this.completed) return;
    if (!this.pagination || !this.summary) {
      throw new Error('DOC PageSource 缺少完成分页所需的元数据');
    }
    if (warnings?.length) {
      this.summary = { ...this.summary, warnings: [...warnings] };
    }
    for (const page of this.pagination.append([], true)) {
      await this.publishPage(page);
    }
    this.completed = true;
    this.writableOutline.complete();
    this.stats.setEstimatedPageCount(this.snapshot.pages.length);
    this.snapshot = {
      ...this.snapshot,
      revision: this.snapshot.revision + 1,
      pageCount: this.snapshot.pages.length,
    };
    this.emitChange();
  }

  hasRenderableContent() {
    return Boolean(this.summary && this.snapshot.pages.length);
  }

  getSummary() {
    if (!this.summary) throw new Error('DOC PageSource 尚无文档摘要');
    return this.summary;
  }

  getOutlineItems() {
    return [...this.outlineItemsById.values()];
  }

  getPerformanceProfile() {
    const profile = createWordPerformanceProfile(this.stats.getSnapshot());
    // 只有提前判定为大文件的会话才创建该 Source，因此页面路径固定使用窗口模式。
    return { ...profile, pageMode: 'windowed' as const };
  }

  getSnapshot = () => this.snapshot;

  subscribe = (listener: () => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  async getPage(index: number, signal?: AbortSignal) {
    await this.waitForPage(index, signal);
    const meta = this.snapshot.pages[index];
    return meta ? this.pageStore.get(meta, signal) : undefined;
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

  retry() {
    // 页面在写入 Store 前已经完成解析；冷层失败时 getPage 会继续走热层。
  }

  dispose() {
    if (this.disposePromise) return this.disposePromise;
    this.completed = true;
    this.writableOutline.complete();
    // 先唤醒等待页面或块定位的调用，再清空订阅，避免主动释放时遗留悬空 Promise。
    this.emitChange();
    this.listeners.clear();
    this.blockPageIndex.clear();
    this.outlineItemsById.clear();
    this.disposePromise = (async () => {
      await this.pageStore.dispose();
      this.resources.dispose();
    })();
    return this.disposePromise;
  }

  private async publishPage(page: PaginatedDocPage) {
    const index = this.snapshot.pages.length;
    const sourceBlockIds = [
      ...new Set(
        page.blocks.flatMap((block) => [
          block.id,
          ...(block.sourceBlockId ? [block.sourceBlockId] : []),
        ]),
      ),
    ];
    const meta: WordPageMeta = {
      id: page.id,
      index,
      revision: 1,
      status: 'ready',
      estimatedContentHeight: this.summary?.page.minHeight ?? 1123,
      sourceBlockIds,
    };
    await this.pageStore.put(meta, page);
    sourceBlockIds.forEach((blockId) =>
      this.blockPageIndex.set(blockId, index),
    );
    this.snapshot = {
      revision: this.snapshot.revision + 1,
      pages: [...this.snapshot.pages, meta],
    };
    this.emitChange();
  }

  private appendOutline(items: readonly WordOutlineItem[]) {
    const nextItems = items.filter((item) => {
      if (this.outlineItemsById.has(item.id)) return false;
      this.outlineItemsById.set(item.id, item);
      return true;
    });
    if (nextItems.length) this.writableOutline.append(nextItems);
  }

  private async waitForPage(index: number, signal?: AbortSignal) {
    await this.waitUntil(
      () => Boolean(this.snapshot.pages[index]) || this.completed,
      signal,
    );
  }

  private async waitUntil(predicate: () => boolean, signal?: AbortSignal) {
    throwIfAborted(signal);
    throwIfAborted(this.signal);
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
        const error = new Error('操作已取消');
        error.name = 'AbortError';
        reject(error);
      };
      unsubscribe = this.subscribe(resolveWhenReady);
      signal?.addEventListener('abort', abort, { once: true });
      this.signal?.addEventListener('abort', abort, { once: true });
      // 订阅与监听安装期间也可能发生取消，安装完成后必须补做一次只读检查。
      if (signal?.aborted || this.signal?.aborted) abort();
    });
  }

  private emitChange() {
    this.listeners.forEach((listener) => listener());
  }
}
