import type {
  DocxMeasuredBlock,
  DocxMeasurementBatch,
} from '../../../docx/docxPagination';
import type {
  DocxPagePreviewSource,
  DocxWordPreviewSummary,
} from '../../../docx/DocxWordPageSource';
import type { DocxPageContent } from '../../../docx/types';
import type { OfficeSearchProvider } from '../../../search/types';
import { createProgressiveWordOutlineProvider } from '../../../word/createMemoryWordOutlineProvider';
import type { WordPerformanceProfile } from '../../../word/types';
import type { WordPageSourceSnapshot } from '../../../word/WordPageSource';
import type { WorkerWordSourceState } from '../../protocol/messages';
import { WorkerSourceClient } from './WorkerSourceClient';

/** 将 Worker 中长期持有的 DOCX Source 适配为现有 Viewer 接口。 */
export class WorkerWordPageSource implements DocxPagePreviewSource {
  readonly pages = this;
  readonly outline = createProgressiveWordOutlineProvider();
  readonly searchProvider: OfficeSearchProvider = {
    kind: 'word',
    search: (query, emit, signal) => this.client.search(query, emit, signal),
  };
  private readonly listeners = new Set<() => void>();
  private readonly unsubscribeUpdate: () => void;
  private readonly unsubscribeFailure: () => void;
  private snapshot: WordPageSourceSnapshot;
  private readonly summary: DocxWordPreviewSummary;
  private outlineItems: WorkerWordSourceState['outlineItems'];
  private performance: WordPerformanceProfile;
  private measurementBatch?: DocxMeasurementBatch;
  private disposed = false;

  constructor(
    private readonly client: WorkerSourceClient,
    initial: WorkerWordSourceState,
  ) {
    this.snapshot = initial.snapshot;
    this.summary = initial.summary as DocxWordPreviewSummary;
    this.outlineItems = initial.outlineItems;
    this.performance = initial.performance;
    this.measurementBatch = initial.measurementBatch as
      | DocxMeasurementBatch
      | undefined;
    this.outline.append(initial.outlineItems);
    if (initial.outlineComplete) this.outline.complete();
    this.unsubscribeUpdate = client.subscribe((source) => {
      if (source.kind !== 'docx') return;
      this.applyState(source);
    });
    this.unsubscribeFailure = client.subscribeFailure((error) => {
      this.snapshot = {
        ...this.snapshot,
        revision: this.snapshot.revision + 1,
        pages: this.snapshot.pages.map((page) =>
          page.status === 'estimated'
            ? {
                ...page,
                revision: page.revision + 1,
                status: 'error',
                errorMessage: error.message,
              }
            : page,
        ),
      };
      this.emitChange();
    });
  }

  getSnapshot = () => this.snapshot;

  subscribe = (listener: () => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  getSummary() {
    return this.summary;
  }

  getPerformanceProfile() {
    return this.performance;
  }

  getOutlineItems() {
    return [...this.outlineItems];
  }

  getMeasurementBatch() {
    return this.measurementBatch;
  }

  hasRenderableContent() {
    return true;
  }

  getPage(index: number, signal?: AbortSignal) {
    return this.client.request<DocxPageContent | undefined>(
      'get-page',
      { index },
      { signal },
    );
  }

  ensureRange(start: number, end: number, signal?: AbortSignal) {
    return this.client.request<void>(
      'ensure-word-range',
      { start, end },
      { signal },
    );
  }

  prioritizeBlock(blockId: string, signal?: AbortSignal) {
    return this.client.request<number>(
      'prioritize-block',
      { blockId },
      { signal },
    );
  }

  retainRange(start: number, end: number) {
    return this.client.retain('retain-word-range', { start, end });
  }

  retry(index: number) {
    void this.client
      .request('retry-word-page', { index })
      .catch(() => undefined);
  }

  commitMeasurement(
    batch: DocxMeasurementBatch,
    measurements: readonly DocxMeasuredBlock[],
    durationMs: number,
  ) {
    return this.client.request<void>('commit-word-measurement', {
      batchId: batch.id,
      measurements,
      durationMs,
    });
  }

  failMeasurement(batch: DocxMeasurementBatch, error: unknown) {
    const message = error instanceof Error ? error.message : String(error);
    void this.client
      .request('fail-word-measurement', { batchId: batch.id, message })
      .catch(() => undefined);
  }

  async dispose() {
    if (this.disposed) return;
    this.disposed = true;
    this.unsubscribeUpdate();
    this.unsubscribeFailure();
    this.listeners.clear();
    await this.client.dispose();
  }

  private applyState(source: WorkerWordSourceState) {
    this.snapshot = source.snapshot;
    this.performance = source.performance;
    this.measurementBatch = source.measurementBatch as
      | DocxMeasurementBatch
      | undefined;
    Object.assign(this.summary, source.summary as DocxWordPreviewSummary);
    this.outlineItems = source.outlineItems;
    this.outline.append(source.outlineItems);
    if (source.outlineComplete) this.outline.complete();
    this.emitChange();
  }

  private emitChange() {
    this.listeners.forEach((listener) => listener());
  }
}
