import { createContentStore, type OfficeContentStore } from '../content-store';
import type { WordPageMeta } from './WordPageSource';

/** Word 页面缓存的容量和回收选项。 */
export type WordPageStoreOptions<TPage> = {
  /** 当前解析或预览会话的标识。 */
  sessionId: string;
  /** 内存热缓存允许占用的最大字节数。 */
  maxMemoryBytes?: number;
  /** 估算单个缓存值占用的字节数。 */
  estimateSize(page: TPage): number;
  /** 在缓存缺页时重新构建指定页面。 */
  recreatePage?(
    index: number,
    signal?: AbortSignal,
  ): Promise<TPage | undefined>;
  /** 报告不会阻断预览的解析或缓存警告。 */
  onWarning?(error: unknown): void;
};

/** Word 页面缓存默认允许占用的内存大小，单位为字节。 */
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
