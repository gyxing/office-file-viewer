import type { RefObject } from 'react';
import { useCallback, useEffect, useRef, useState } from 'react';
import type { WordOutlineItem } from '../../services/word/types';

type UseWordOutlineNavigationOptions = {
  /** 源文档大纲条目。 */
  items: WordOutlineItem[];
  /** Word 正文唯一的滚动容器。 */
  scrollContainerRef: RefObject<HTMLElement>;
  /** 缩放、分页等会改变目标位置的布局键。 */
  layoutKey: string;
  /** 侧栏展开时才启用观察和滚动同步。 */
  enabled: boolean;
};

type OutlineTarget = {
  /** 对应大纲条目的稳定键。 */
  key: string;
  /** 目标相对滚动内容顶部的位置。 */
  top: number;
  /** 正文段落元素。 */
  element: HTMLElement;
};

/** 使用二分查找定位滚动标记线之前最后一个大纲条目。 */
function findActiveTarget(targets: OutlineTarget[], marker: number) {
  let low = 0;
  let high = targets.length - 1;
  let activeIndex = 0;
  while (low <= high) {
    const middle = Math.floor((low + high) / 2);
    if (targets[middle].top <= marker) {
      activeIndex = middle;
      low = middle + 1;
    } else {
      high = middle - 1;
    }
  }
  return targets[activeIndex];
}

/** 管理 Word 大纲定位、滚动高亮和布局缓存。 */
export function useWordOutlineNavigation({
  items,
  scrollContainerRef,
  layoutKey,
  enabled,
}: UseWordOutlineNavigationOptions) {
  const [activeKey, setActiveKey] = useState<string>();
  const targetsRef = useRef<OutlineTarget[]>([]);
  const frameRef = useRef<number>();
  const usesInternalScrollRef = useRef(true);

  const commitActiveKey = useCallback((key: string | undefined) => {
    setActiveKey((current) => (current === key ? current : key));
  }, []);

  const syncFromScroll = useCallback(() => {
    const scroller = scrollContainerRef.current;
    const targets = targetsRef.current;
    if (!scroller || !targets.length) return;
    const marker = usesInternalScrollRef.current
      ? scroller.scrollTop + 32
      : window.scrollY + 32;
    const active = findActiveTarget(targets, marker);
    commitActiveKey(active?.key);
  }, [commitActiveKey, scrollContainerRef]);

  const scheduleScrollSync = useCallback(() => {
    if (frameRef.current !== undefined) return;
    frameRef.current = requestAnimationFrame(() => {
      frameRef.current = undefined;
      syncFromScroll();
    });
  }, [syncFromScroll]);

  useEffect(() => {
    const scroller = scrollContainerRef.current;
    if (!enabled || !scroller || !items.length) {
      targetsRef.current = [];
      commitActiveKey(undefined);
      return;
    }

    const itemByTargetId = new Map(
      items.map((item) => [item.targetBlockId, item]),
    );
    const usesInternalScroll =
      scroller.scrollHeight > scroller.clientHeight + 1;
    usesInternalScrollRef.current = usesInternalScroll;
    const scrollEventTarget: HTMLElement | Window = usesInternalScroll
      ? scroller
      : window;
    const layoutFrame = requestAnimationFrame(() => {
      const scrollerRect = scroller.getBoundingClientRect();
      targetsRef.current = Array.from(
        scroller.querySelectorAll<HTMLElement>(
          '[data-office-word-outline-target]',
        ),
      )
        .flatMap((element) => {
          const targetId = element.dataset.officeWordOutlineTarget;
          const item = targetId ? itemByTargetId.get(targetId) : undefined;
          if (!item) return [];
          const rect = element.getBoundingClientRect();
          return [
            {
              key: item.id,
              top: usesInternalScroll
                ? scroller.scrollTop + rect.top - scrollerRect.top
                : window.scrollY + rect.top,
              element,
            },
          ];
        })
        .sort((left, right) => left.top - right.top);
      syncFromScroll();
    });

    scrollEventTarget.addEventListener('scroll', scheduleScrollSync, {
      passive: true,
    });
    const observer =
      typeof IntersectionObserver === 'undefined'
        ? undefined
        : new IntersectionObserver(
            (entries) => {
              const visible = entries
                .filter((entry) => entry.isIntersecting)
                .sort(
                  (left, right) =>
                    Math.abs(left.boundingClientRect.top) -
                    Math.abs(right.boundingClientRect.top),
                )[0];
              const targetId = (visible?.target as HTMLElement | undefined)
                ?.dataset.officeWordOutlineTarget;
              const item = targetId ? itemByTargetId.get(targetId) : undefined;
              if (item) commitActiveKey(item.id);
            },
            {
              root: usesInternalScroll ? scroller : null,
              rootMargin: '0px 0px -70% 0px',
              threshold: 0,
            },
          );

    const observeFrame = requestAnimationFrame(() => {
      targetsRef.current.forEach((target) => observer?.observe(target.element));
    });
    return () => {
      cancelAnimationFrame(layoutFrame);
      cancelAnimationFrame(observeFrame);
      if (frameRef.current !== undefined) {
        cancelAnimationFrame(frameRef.current);
        frameRef.current = undefined;
      }
      scrollEventTarget.removeEventListener('scroll', scheduleScrollSync);
      observer?.disconnect();
      targetsRef.current = [];
    };
  }, [
    commitActiveKey,
    enabled,
    items,
    layoutKey,
    scheduleScrollSync,
    scrollContainerRef,
    syncFromScroll,
  ]);

  const selectTarget = useCallback(
    (item: WordOutlineItem) => {
      const scroller = scrollContainerRef.current;
      const target = targetsRef.current.find(
        (candidate) => candidate.key === item.id,
      );
      if (!scroller || !target) return;
      commitActiveKey(item.id);
      // 遵循系统减少动态效果设置，避免强制平滑滚动引发不适。
      const behavior = window.matchMedia?.('(prefers-reduced-motion: reduce)')
        .matches
        ? 'auto'
        : 'smooth';
      if (usesInternalScrollRef.current) {
        scroller.scrollTo({
          top: Math.max(0, target.top - 20),
          behavior,
        });
      } else {
        target.element.scrollIntoView({ behavior, block: 'start' });
      }
    },
    [commitActiveKey, scrollContainerRef],
  );

  return { activeKey, selectTarget };
}
