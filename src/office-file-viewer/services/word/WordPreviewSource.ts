import type { WordOutlineProvider } from './WordOutlineProvider';
import type { WordPageSource } from './WordPageSource';
import type { WordPerformanceProfile } from './types';

/** Word Viewer 消费的页面、大纲和性能画像组合来源。 */
export interface WordPreviewSource<TPage> {
  readonly pages: WordPageSource<TPage>;
  readonly outline: WordOutlineProvider;
  getPerformanceProfile(): WordPerformanceProfile;
  dispose(): Promise<void>;
}
