/** 当前只读查看器支持的页级切换类型。 */
export type PresentationTransitionType =
  | 'fade'
  | 'push'
  | 'wipe'
  | 'split'
  | 'cover'
  | 'uncover';

/** 页级切换使用的方向或分割方式。 */
export type PresentationTransitionDirection =
  | 'left'
  | 'right'
  | 'up'
  | 'down'
  | 'horizontal-in'
  | 'horizontal-out'
  | 'vertical-in'
  | 'vertical-out';

/** 从源文件恢复的单页切换设置。 */
export type PresentationTransition = Readonly<{
  /** 浏览器能够还原的切换类型。 */
  type: PresentationTransitionType;
  /** 源切换声明的运动方向。 */
  direction?: PresentationTransitionDirection;
  /** 已限制在 100 至 2000 毫秒之间的持续时间。 */
  durationMs: number;
}>;

/** 是否按源文件启用页级切换；默认 false。 */
export type OfficeFileViewerPresentationTransitions = false | 'source';

/** 工具栏上一页/下一页产生的单次切换意图。 */
export type PresentationNavigationIntent = Readonly<{
  /** 每次工具栏翻页递增的标识，用于取消和去重。 */
  token: number;
  /** 目标幻灯片的零基索引。 */
  targetIndex: number;
  /** 用户请求的翻页方向。 */
  direction: 'previous' | 'next';
}>;

/** 把不可信的源持续时间约束到可交互范围。 */
export function normalizePresentationTransitionDuration(value: number) {
  if (!Number.isFinite(value)) return 500;
  return Math.min(2000, Math.max(100, Math.round(value)));
}
