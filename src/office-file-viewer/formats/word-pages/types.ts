/** 当前页面窗口的首尾索引和两端占位高度。 */
export type WordPageWindowRange = {
  /** 当前范围的起始位置。 */
  start: number;
  /** 当前范围的结束位置。 */
  end: number;
  /** 虚拟列表顶部占位高度，单位为标准化渲染像素。 */
  topSpacerHeight: number;
  /** 虚拟列表底部占位高度，单位为标准化渲染像素。 */
  bottomSpacerHeight: number;
};

/** 页面窗口内部的异步读取状态。 */
export type WordPageLoadState<TPage> =
  | {
      /** 当前页面正在加载。 */
      status: 'loading';
    }
  | {
      /** 当前加载或解析状态。 */
      status: 'ready';
      /** 当前关联的页面模型。 */
      page: TPage;
      /** 数据源变更时递增的修订号。 */
      revision: number;
    }
  | {
      /** 当前加载或解析状态。 */
      status: 'error';
      /** 当前操作产生的错误；未提供表示没有错误。 */
      error?: unknown;
      /** 数据源变更时递增的修订号。 */
      revision: number;
    };

/** 页面窗口向大纲导航暴露的最小控制面。 */
export interface WordPageNavigationController {
  /** 将指定页面滚动到预览区域。 */
  scrollToPage(index: number, offset?: number): void;
  /** 确保指定页面已经挂载到可滚动容器。 */
  ensurePageMounted(index: number, signal?: AbortSignal): Promise<HTMLElement>;
  /** 返回当前已经挂载的页面索引范围。 */
  getMountedRange(): {
    /** 当前范围的起始位置。 */
    start: number;
    /** 当前范围的结束位置。 */
    end: number;
  };
}
