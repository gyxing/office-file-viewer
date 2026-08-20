import type { WordRevisionKind } from '../../services/word/review/types';

/** Word 页侧标记的语义类别。 */
export type WordMarkupCalloutType = 'comment' | 'revision';

/** 当前视口内一处批注或修订的正文锚点。 */
export type WordMarkupCalloutAnchor = Readonly<{
  /** 合并语义类别与源标识后的稳定键。 */
  key: string;
  /** 源批注或修订标识。 */
  id: string;
  /** 当前标记属于批注或修订。 */
  type: WordMarkupCalloutType;
  /** 修订标记的插入、删除、移动或格式类别。 */
  revisionKind?: WordRevisionKind;
  /** 源文件中的作者。 */
  author?: string;
  /** 源文件中的时间。 */
  createdAt?: string;
  /** 批注正文或修订摘要。 */
  excerpt: string;
  /** 批注线程是否已经解决。 */
  resolved?: boolean;
  /** 正文锚点左边界。 */
  left: number;
  /** 正文锚点上边界。 */
  top: number;
  /** 正文锚点右边界。 */
  right: number;
  /** 正文锚点下边界。 */
  bottom: number;
  /** 当前标记内容需要的稳定高度。 */
  height: number;
}>;

/** 当前页面与滚动视口共同决定的纵向标记区域。 */
export type WordMarkupCalloutBoundary = Readonly<{
  /** 当前页面可见区域上边界。 */
  top: number;
  /** 当前页面可见区域下边界。 */
  bottom: number;
  /** Word 页面右边界。 */
  pageRight: number;
}>;

/** 一处 Word 页侧标记的最终布局。 */
export type WordMarkupCalloutLayout = WordMarkupCalloutAnchor &
  Readonly<{
    /** 页侧存在多列标记时采用的零基轨道索引。 */
    laneIndex: number;
    /** 标记内容在滚动坐标系中的横坐标。 */
    calloutLeft: number;
    /** 标记内容在滚动坐标系中的纵坐标。 */
    calloutTop: number;
    /** 标记内容宽度。 */
    width: number;
    /** 连接线指向正文的横坐标。 */
    anchorX: number;
    /** 连接线指向正文的纵坐标。 */
    anchorY: number;
    /** 连接线抵达页侧标记竖线时的纵坐标。 */
    connectorY: number;
  }>;

/** 页面正文与页侧标记区之间的距离。 */
export const WORD_MARKUP_RAIL_GAP = 16;
/** 单列页侧标记的宽度。 */
export const WORD_MARKUP_CALLOUT_WIDTH = 300;
/** 页面与一列页侧标记组成的基础标记区宽度。 */
export const WORD_MARKUP_BASE_RAIL_WIDTH =
  WORD_MARKUP_RAIL_GAP + WORD_MARKUP_CALLOUT_WIDTH;
/** 多列页侧标记之间的距离。 */
const WORD_MARKUP_LANE_GAP = 16;
/** 同列相邻标记之间的距离。 */
const WORD_MARKUP_STACK_GAP = 8;
/** 标记与当前可见区域上下边界之间的距离。 */
const WORD_MARKUP_VIEWPORT_INSET = 8;

/** 将动态坐标约束到当前页面可见区域。 */
function clamp(value: number, minimum: number, maximum: number) {
  return Math.min(Math.max(value, minimum), Math.max(minimum, maximum));
}

/** 在同一页侧轨道中按正文顺序纵向避让。 */
function packLane(
  callouts: readonly WordMarkupCalloutLayout[],
  boundary: WordMarkupCalloutBoundary,
) {
  const sorted = [...callouts].sort(
    (first, second) => first.anchorY - second.anchorY,
  );
  let cursor = boundary.top + WORD_MARKUP_VIEWPORT_INSET;
  const forward = sorted.map((callout) => {
    const calloutTop = clamp(
      Math.max(callout.calloutTop, cursor),
      boundary.top + WORD_MARKUP_VIEWPORT_INSET,
      boundary.bottom - callout.height - WORD_MARKUP_VIEWPORT_INSET,
    );
    cursor = calloutTop + callout.height + WORD_MARKUP_STACK_GAP;
    return {
      ...callout,
      calloutTop,
      connectorY:
        calloutTop +
        Math.min(callout.height / 2, callout.type === 'comment' ? 24 : 18),
    };
  });
  cursor = boundary.bottom - WORD_MARKUP_VIEWPORT_INSET;
  for (let index = forward.length - 1; index >= 0; index -= 1) {
    const callout = forward[index];
    const calloutTop = Math.max(
      boundary.top + WORD_MARKUP_VIEWPORT_INSET,
      Math.min(callout.calloutTop, cursor - callout.height),
    );
    forward[index] = {
      ...callout,
      calloutTop,
      connectorY:
        calloutTop +
        Math.min(callout.height / 2, callout.type === 'comment' ? 24 : 18),
    };
    cursor = calloutTop - WORD_MARKUP_STACK_GAP;
  }
  return forward;
}

/** 把可见批注和修订统一排入 Word 页面右侧标记区。 */
export function layoutWordMarkupCallouts(
  anchors: readonly WordMarkupCalloutAnchor[],
  boundary: WordMarkupCalloutBoundary,
) {
  if (!anchors.length) return [];
  const maximumHeight = Math.max(...anchors.map((anchor) => anchor.height));
  const availableHeight = Math.max(
    1,
    boundary.bottom - boundary.top - WORD_MARKUP_VIEWPORT_INSET * 2,
  );
  const laneCapacity = Math.max(
    1,
    Math.floor(
      (availableHeight + WORD_MARKUP_STACK_GAP) /
        (maximumHeight + WORD_MARKUP_STACK_GAP),
    ),
  );
  const sorted = [...anchors].sort(
    (first, second) =>
      (first.top + first.bottom) / 2 - (second.top + second.bottom) / 2,
  );
  const drafts = sorted.map<WordMarkupCalloutLayout>((anchor, index) => {
    const laneIndex = Math.floor(index / laneCapacity);
    const calloutLeft =
      boundary.pageRight +
      WORD_MARKUP_RAIL_GAP +
      laneIndex * (WORD_MARKUP_CALLOUT_WIDTH + WORD_MARKUP_LANE_GAP);
    const anchorY = (anchor.top + anchor.bottom) / 2;
    const calloutTop = clamp(
      anchorY - anchor.height / 2,
      boundary.top + WORD_MARKUP_VIEWPORT_INSET,
      boundary.bottom - anchor.height - WORD_MARKUP_VIEWPORT_INSET,
    );
    return {
      ...anchor,
      laneIndex,
      calloutLeft,
      calloutTop,
      width: WORD_MARKUP_CALLOUT_WIDTH,
      anchorX: anchor.right,
      anchorY,
      connectorY:
        calloutTop +
        Math.min(anchor.height / 2, anchor.type === 'comment' ? 24 : 18),
    };
  });
  const lanes = [...new Set(drafts.map((callout) => callout.laneIndex))];
  return lanes.flatMap((laneIndex) =>
    packLane(
      drafts.filter((callout) => callout.laneIndex === laneIndex),
      boundary,
    ),
  );
}

/** 返回当前页侧标记区需要占用的最右滚动坐标。 */
export function getWordMarkupRailExtent(
  pageRight: number,
  callouts: readonly WordMarkupCalloutLayout[],
) {
  const baseExtent =
    pageRight + WORD_MARKUP_RAIL_GAP + WORD_MARKUP_CALLOUT_WIDTH;
  return callouts.reduce(
    (maximum, callout) =>
      Math.max(maximum, callout.calloutLeft + callout.width),
    baseExtent,
  );
}
