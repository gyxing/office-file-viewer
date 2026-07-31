import type { WordOutlineProvider } from './WordOutlineProvider';
import type { WordPageSource } from './WordPageSource';
import type { WordPerformanceProfile } from './types';

/** Word Viewer 消费的页面、大纲和性能画像组合来源。 */
export interface WordPreviewSource<TPage> {
  /** 按文档顺序排列的页面。 */
  readonly pages: WordPageSource<TPage>;
  /** Word 大纲相关文案。 */
  readonly outline: WordOutlineProvider;
  /** 返回当前文档采用的性能配置。 */
  getPerformanceProfile(): WordPerformanceProfile;
  /** 幂等释放当前对象持有的资源和订阅。 */
  dispose(): Promise<void>;
}
