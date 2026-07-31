import type { RefObject } from 'react';
import { useEffect, useState } from 'react';
import type { PageHeightIndex } from './PageHeightIndex';
import type { WordPageWindowRange } from './types';

/** Word 虚拟列表始终保留挂载的最少页面数。 */
const MINIMUM_MOUNTED_PAGES = 3;
/** Word 虚拟列表在可视区外预渲染的视口倍数。 */
const OVERSCAN_VIEWPORTS = 2;

function createRange(
  heightIndex: PageHeightIndex,
  scrollTop: number,
  viewportHeight: number,
): WordPageWindowRange {
  const pageCount = heightIndex.length;
  if (!pageCount) {
    return {
      start: 0,
      end: 0,
      topSpacerHeight: 0,
      bottomSpacerHeight: 0,
    };
  }
  const overscan = Math.max(0, viewportHeight) * OVERSCAN_VIEWPORTS;
  let start = heightIndex.findIndexAtOffset(Math.max(0, scrollTop - overscan));
  let end =
    heightIndex.findIndexAtOffset(
      Math.min(heightIndex.total(), scrollTop + viewportHeight + overscan),
    ) + 1;
  const missing = MINIMUM_MOUNTED_PAGES - (end - start);
  if (missing > 0) {
    const before = Math.min(start, Math.ceil(missing / 2));
    start -= before;
    end = Math.min(pageCount, end + missing - before);
    start = Math.max(0, end - MINIMUM_MOUNTED_PAGES);
  }
  return {
    start,
    end,
    topSpacerHeight: heightIndex.prefix(start),
    bottomSpacerHeight: heightIndex.total() - heightIndex.prefix(end),
  };
}

/** 合并滚动事件，并按上下约两个视口计算有界页面窗口。 */
export function useWordPageWindow(
  scrollerRef: RefObject<HTMLElement>,
  heightIndex: PageHeightIndex,
  layoutRevision: string,
  heightRevision: number,
) {
  const [range, setRange] = useState(() => createRange(heightIndex, 0, 0));

  useEffect(() => {
    const scroller = scrollerRef.current;
    if (!scroller) return;
    const usesInternalScroll =
      scroller.scrollHeight > scroller.clientHeight + 1;
    const scrollTarget: HTMLElement | Window = usesInternalScroll
      ? scroller
      : window;
    let frame: number | undefined;
    const update = () => {
      frame = undefined;
      const contentOffset = usesInternalScroll
        ? scroller.scrollTop
        : Math.max(0, -scroller.getBoundingClientRect().top);
      const viewportHeight = usesInternalScroll
        ? scroller.clientHeight
        : window.innerHeight;
      setRange(createRange(heightIndex, contentOffset, viewportHeight));
    };
    const schedule = () => {
      if (frame !== undefined) return;
      frame = requestAnimationFrame(update);
    };
    update();
    scrollTarget.addEventListener('scroll', schedule, { passive: true });
    window.addEventListener('resize', schedule, { passive: true });
    return () => {
      if (frame !== undefined) cancelAnimationFrame(frame);
      scrollTarget.removeEventListener('scroll', schedule);
      window.removeEventListener('resize', schedule);
    };
  }, [heightIndex, heightRevision, layoutRevision, scrollerRef]);

  return range;
}
