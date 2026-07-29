import { createContentStore, type OfficeContentStore } from '../content-store';
import type { WordPageMeta } from './WordPageSource';

export type WordPageStoreOptions<TPage> = {
  sessionId: string;
  maxMemoryBytes?: number;
  estimateSize(page: TPage): number;
  recreatePage?(
    index: number,
    signal?: AbortSignal,
  ): Promise<TPage | undefined>;
  onWarning?(error: unknown): void;
};

const DEFAULT_WORD_PAGE_MEMORY_BYTES = 32 * 1024 * 1024;

/** 以内存热层和 IndexedDB 冷层保存大文件页面正文。 */
export class WordPageStore<TPage> {
  private readonly store: OfficeContentStore<WordPageMeta, TPage>;
  private readonly recreatePage?: WordPageStoreOptions<TPage>['recreatePage'];

  constructor(options: WordPageStoreOptions<TPage>) {
    this.recreatePage = options.recreatePage;
    this.store = createContentStore({
      sessionId: options.sessionId,
      namespace: 'word-pages',
      maxMemoryBytes: options.maxMemoryBytes ?? DEFAULT_WORD_PAGE_MEMORY_BYTES,
      estimateSize: options.estimateSize,
      onWarning: options.onWarning,
    });
  }

  /** 写入递增 revision 的页面和轻量元数据。 */
  async put(meta: WordPageMeta, page: TPage) {
    await this.store.put({
      key: meta.id,
      revision: meta.revision,
      meta,
      value: page,
      updatedAt: Date.now(),
    });
  }

  /** 优先从热层/冷层读取，冷层不可用时允许解析器重建页面。 */
  async get(meta: WordPageMeta, signal?: AbortSignal) {
    const stored = await this.store.get(meta.id, signal);
    if (stored?.value !== undefined) return stored.value;
    return this.recreatePage?.(meta.index, signal);
  }

  /** 固定当前窗口页面，避免 LRU 驱逐正在显示的正文。 */
  retain(metas: readonly WordPageMeta[]) {
    return this.store.pin(metas.map((meta) => meta.id));
  }

  delete(meta: WordPageMeta) {
    return this.store.delete(meta.id);
  }

  dispose() {
    return this.store.dispose();
  }
}
