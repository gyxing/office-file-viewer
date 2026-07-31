import { useEffect, useMemo, useState, type RefObject } from 'react';

/** 缩略图虚拟列表当前挂载的索引范围。 */
export type PresentationThumbnailRange = {
  /** 当前范围的起始位置。 */
  start: number;
  /** 当前范围的结束位置。 */
  end: number;
  /** 虚拟列表顶部占位区域。 */
  topSpacer: number;
  /** 虚拟列表底部占位区域。 */
  bottomSpacer: number;
};

/** 缩略图虚拟列表在可视区两侧额外渲染的项目数。 */
const THUMBNAIL_OVERSCAN = 5;

/** 根据缩略图侧栏自身尺寸计算虚拟窗口，不向组件公共 API 暴露高度参数。 */
export function usePresentationThumbnailWindow(
  viewportRef: RefObject<HTMLElement>,
  itemCount: number,
  itemHeight: number,
) {
  const [viewport, setViewport] = useState({ height: 0, scrollTop: 0 });

  useEffect(() => {
    const element = viewportRef.current;
    if (!element) return undefined;

    const updateSize = () => {
      setViewport((current) => ({
        height: element.clientHeight,
        scrollTop: current.scrollTop,
      }));
    };
    const updateScroll = () => {
      setViewport((current) => ({
        height: current.height || element.clientHeight,
        scrollTop: element.scrollTop,
      }));
    };

    updateSize();
    element.addEventListener('scroll', updateScroll, { passive: true });
    const observer =
      typeof ResizeObserver === 'undefined'
        ? undefined
        : new ResizeObserver(updateSize);
    observer?.observe(element);
    return () => {
      element.removeEventListener('scroll', updateScroll);
      observer?.disconnect();
    };
  }, [viewportRef]);

  return useMemo<PresentationThumbnailRange>(() => {
    const visibleStart = Math.floor(viewport.scrollTop / itemHeight);
    const visibleEnd = Math.ceil(
      (viewport.scrollTop + Math.max(viewport.height, itemHeight)) / itemHeight,
    );
    const start = Math.max(0, visibleStart - THUMBNAIL_OVERSCAN);
    const end = Math.min(itemCount, visibleEnd + THUMBNAIL_OVERSCAN);
    return {
      start,
      end,
      topSpacer: start * itemHeight,
      bottomSpacer: Math.max(0, (itemCount - end) * itemHeight),
    };
  }, [itemCount, itemHeight, viewport.height, viewport.scrollTop]);
}
