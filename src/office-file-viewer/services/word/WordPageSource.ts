/** Word 页面常驻内存的轻量状态，不包含正文 blocks。 */
export type WordPageMeta = {
  id: string;
  index: number;
  revision: number;
  status: 'estimated' | 'ready' | 'error';
  estimatedContentHeight: number;
  sourceBlockIds: readonly string[];
  errorMessage?: string;
};

/** 页面 Source 的版本化快照。 */
export type WordPageSourceSnapshot = {
  revision: number;
  pageCount?: number;
  pages: readonly WordPageMeta[];
};

/** 为普通数组和渐进分页统一提供按范围页面读取能力。 */
export interface WordPageSource<TPage> {
  getSnapshot(): WordPageSourceSnapshot;
  subscribe(listener: () => void): () => void;
  getPage(index: number, signal?: AbortSignal): Promise<TPage | undefined>;
  ensureRange(start: number, end: number, signal?: AbortSignal): Promise<void>;
  prioritizeBlock(blockId: string, signal?: AbortSignal): Promise<number>;
  retainRange(start: number, end: number): () => void;
  retry(index: number): void;
  dispose(): Promise<void>;
}
