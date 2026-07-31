/** Word 页面常驻内存的轻量状态，不包含正文 blocks。 */
export type WordPageMeta = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 在所属集合中的零基索引。 */
  index: number;
  /** 数据源变更时递增的修订号。 */
  revision: number;
  /** 当前加载或解析状态。 */
  status: 'estimated' | 'ready' | 'error';
  /** 页面内容区域的估算高度。 */
  estimatedContentHeight: number;
  /** 组成当前页面的源内容块标识。 */
  sourceBlockIds: readonly string[];
  /** 当前错误对应的可读说明。 */
  errorMessage?: string;
};

/** 页面 Source 的版本化快照。 */
export type WordPageSourceSnapshot = {
  /** 数据源变更时递增的修订号。 */
  revision: number;
  /** 页面数量。 */
  pageCount?: number;
  /** 按文档顺序排列的页面。 */
  pages: readonly WordPageMeta[];
};

/** 为普通数组和渐进分页统一提供按范围页面读取能力。 */
export interface WordPageSource<TPage> {
  /** 返回当前可观察状态的只读快照。 */
  getSnapshot(): WordPageSourceSnapshot;
  /** 订阅状态快照变化，并返回取消订阅函数。 */
  subscribe(listener: () => void): () => void;
  /** 读取指定索引的页面模型。 */
  getPage(index: number, signal?: AbortSignal): Promise<TPage | undefined>;
  /** 确保指定内容范围已经开始加载或可用。 */
  ensureRange(start: number, end: number, signal?: AbortSignal): Promise<void>;
  /** 提升包含指定内容块页面的加载优先级。 */
  prioritizeBlock(blockId: string, signal?: AbortSignal): Promise<number>;
  /** 保留指定可视范围并回收远离窗口的缓存内容。 */
  retainRange(start: number, end: number): () => void;
  /** 重新加载此前失败的指定内容。 */
  retry(index: number): void;
  /** 幂等释放当前对象持有的资源和订阅。 */
  dispose(): Promise<void>;
}
