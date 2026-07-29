/** 当前页面窗口的首尾索引和两端占位高度。 */
export type WordPageWindowRange = {
  start: number;
  end: number;
  topSpacerHeight: number;
  bottomSpacerHeight: number;
};

/** 页面窗口内部的异步读取状态。 */
export type WordPageLoadState<TPage> =
  | { status: 'loading' }
  | { status: 'ready'; page: TPage; revision: number }
  | { status: 'error'; error?: unknown; revision: number };

/** 页面窗口向大纲导航暴露的最小控制面。 */
export interface WordPageNavigationController {
  scrollToPage(index: number, offset?: number): void;
  ensurePageMounted(index: number, signal?: AbortSignal): Promise<HTMLElement>;
  getMountedRange(): { start: number; end: number };
}
