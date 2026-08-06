import type { MutableRefObject, RefObject } from 'react';
import { useCallback, useEffect, useRef, useState } from 'react';
import type { WordPageSource } from '../../services/word/WordPageSource';
import type { WordOutlineItem } from '../../services/word/types';
import type { WordBlockPageIndex } from '../word-pages/WordBlockPageIndex';
import type { WordPageNavigationController } from '../word-pages/types';

/** Word 大纲定位 Hook 的输入选项。 */
type UseWordOutlineNavigationOptions = {
  /** 可用于正文定位的大纲条目。 */
  items: WordOutlineItem[];
  /** 滚动容器的可变引用。 */
  scrollContainerRef: RefObject<HTMLElement>;
  /** 页面布局变化时更新的稳定键。 */
  layoutKey: string;
  /** 当前能力是否启用。 */
  enabled: boolean;
  /** 当前使用的页面加载模式。 */
  pageMode: 'normal' | 'windowed';
  /** 按需提供页面内容的数据源。 */
  pageSource?: WordPageSource<unknown>;
  /** 正文块到页面索引的查询结构。 */
  blockPageIndex?: WordBlockPageIndex;
  /** 页面导航控制器的可变引用。 */
  pageNavigationControllerRef?: MutableRefObject<
    WordPageNavigationController | undefined
  >;
  /** 当前文档解析会话的标识。 */
  documentSessionId: string;
};

/** 大纲条目在正文中的定位目标。 */
type OutlineTarget = {
  /** 用于稳定识别或缓存当前项目的键。 */
  key: string;
  /** 定位目标相对页面顶部的位置。 */
  top: number;
  /** 大纲条目对应的正文 DOM 元素。 */
  element: HTMLElement;
};

/** 正文定位时保留少量顶部间距，避免标题紧贴容器边缘。 */
const OUTLINE_TARGET_TOP_OFFSET = 20;
/** 连续稳定两帧后再恢复滚动跟随，避免延迟回调覆盖用户点击项。 */
const PROGRAMMATIC_SCROLL_STABLE_FRAME_COUNT = 2;
/** 浏览器滚动异常时仍需解除高亮锁，避免后续手动滚动失效。 */
const PROGRAMMATIC_SCROLL_MAX_FRAME_COUNT = 120;
/** 亚像素滚动会产生小数误差，位置差异在该范围内视为稳定。 */
const PROGRAMMATIC_SCROLL_POSITION_TOLERANCE = 0.5;

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

function findOutlineElement(
  root: ParentNode,
  targetBlockId: string,
): HTMLElement | undefined {
  return Array.from(
    root.querySelectorAll<HTMLElement>('[data-office-word-outline-target]'),
  ).find(
    (element) => element.dataset.officeWordOutlineTarget === targetBlockId,
  );
}

/** 管理普通 DOM 与页面窗口两种模式的大纲定位和滚动高亮。 */
export function useWordOutlineNavigation({
  items,
  scrollContainerRef,
  layoutKey,
  enabled,
  pageMode,
  pageSource,
  blockPageIndex,
  pageNavigationControllerRef,
  documentSessionId,
}: UseWordOutlineNavigationOptions) {
  const [activeKey, setActiveKey] = useState<string>();
  const targetsRef = useRef<OutlineTarget[]>([]);
  const frameRef = useRef<number>();
  const navigationAbortRef = useRef<AbortController>();
  const navigationSettleFrameRef = useRef<number>();
  const usesInternalScrollRef = useRef(true);
  const programmaticNavigationRef = useRef(false);

  const commitActiveKey = useCallback((key: string | undefined) => {
    setActiveKey((current) => (current === key ? current : key));
  }, []);

  const syncFromScroll = useCallback(() => {
    if (programmaticNavigationRef.current) return;
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

  const cancelNavigationSettle = useCallback(() => {
    if (navigationSettleFrameRef.current === undefined) return;
    cancelAnimationFrame(navigationSettleFrameRef.current);
    navigationSettleFrameRef.current = undefined;
  }, []);

  const settleProgrammaticNavigation = useCallback(
    (element: HTMLElement, expectedPosition: number) => {
      // 平滑滚动完成前保留用户点击项，避免异步滚动回调抢占高亮。
      cancelNavigationSettle();
      const scroller = scrollContainerRef.current;
      if (!scroller) {
        programmaticNavigationRef.current = false;
        return;
      }
      const readPosition = () =>
        usesInternalScrollRef.current ? scroller.scrollTop : window.scrollY;
      const initialPosition = readPosition();
      let previousPosition: number | undefined;
      let stableFrameCount = 0;
      let frameCount = 0;
      let hasMoved = false;

      const checkPosition = () => {
        const currentPosition = readPosition();
        hasMoved ||=
          Math.abs(currentPosition - initialPosition) >
          PROGRAMMATIC_SCROLL_POSITION_TOLERANCE;
        stableFrameCount =
          previousPosition !== undefined &&
          Math.abs(currentPosition - previousPosition) <=
            PROGRAMMATIC_SCROLL_POSITION_TOLERANCE
            ? stableFrameCount + 1
            : 0;
        previousPosition = currentPosition;
        frameCount += 1;
        const reachedExpectedPosition =
          Math.abs(currentPosition - expectedPosition) <=
          PROGRAMMATIC_SCROLL_POSITION_TOLERANCE;
        const hasSettled =
          stableFrameCount >= PROGRAMMATIC_SCROLL_STABLE_FRAME_COUNT &&
          (hasMoved || reachedExpectedPosition);

        if (hasSettled || frameCount >= PROGRAMMATIC_SCROLL_MAX_FRAME_COUNT) {
          navigationSettleFrameRef.current = undefined;
          if (element.isConnected) {
            element.tabIndex = -1;
            element.focus({ preventScroll: true });
          }
          programmaticNavigationRef.current = false;
          // 底部无法贴顶属于合法落点，只有滚动中断时才按当前位置校正。
          if (!reachedExpectedPosition) scheduleScrollSync();
          return;
        }
        navigationSettleFrameRef.current = requestAnimationFrame(checkPosition);
      };

      navigationSettleFrameRef.current = requestAnimationFrame(checkPosition);
    },
    [cancelNavigationSettle, scheduleScrollSync, scrollContainerRef],
  );

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
    let intersectionObserver: IntersectionObserver | undefined;
    let rebuildFrame: number | undefined;

    const rebuildTargets = () => {
      rebuildFrame = undefined;
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
      intersectionObserver?.disconnect();
      targetsRef.current.forEach((target) =>
        intersectionObserver?.observe(target.element),
      );
      syncFromScroll();
    };
    const scheduleRebuild = () => {
      if (rebuildFrame !== undefined) return;
      rebuildFrame = requestAnimationFrame(rebuildTargets);
    };

    intersectionObserver =
      typeof IntersectionObserver === 'undefined'
        ? undefined
        : new IntersectionObserver(
            // 观察器回调只表示可见范围发生变化，最终高亮必须基于完整目标列表计算。
            scheduleScrollSync,
            {
              root: usesInternalScroll ? scroller : null,
              rootMargin: '0px 0px -70% 0px',
              threshold: 0,
            },
          );
    scheduleRebuild();
    scrollEventTarget.addEventListener('scroll', scheduleScrollSync, {
      passive: true,
    });
    const mutationObserver =
      pageMode === 'windowed' && typeof MutationObserver !== 'undefined'
        ? new MutationObserver(scheduleRebuild)
        : undefined;
    mutationObserver?.observe(scroller, { childList: true, subtree: true });

    return () => {
      if (rebuildFrame !== undefined) cancelAnimationFrame(rebuildFrame);
      if (frameRef.current !== undefined) {
        cancelAnimationFrame(frameRef.current);
        frameRef.current = undefined;
      }
      scrollEventTarget.removeEventListener('scroll', scheduleScrollSync);
      mutationObserver?.disconnect();
      intersectionObserver?.disconnect();
      targetsRef.current = [];
    };
  }, [
    commitActiveKey,
    enabled,
    items,
    layoutKey,
    pageMode,
    scheduleScrollSync,
    scrollContainerRef,
    syncFromScroll,
  ]);

  useEffect(
    () => () => {
      navigationAbortRef.current?.abort();
      navigationAbortRef.current = undefined;
      cancelNavigationSettle();
      programmaticNavigationRef.current = false;
    },
    [cancelNavigationSettle, documentSessionId, enabled],
  );

  const scrollToElement = useCallback(
    (element: HTMLElement, item: WordOutlineItem) => {
      const scroller = scrollContainerRef.current;
      if (!scroller) {
        programmaticNavigationRef.current = false;
        return;
      }
      const behavior = window.matchMedia?.('(prefers-reduced-motion: reduce)')
        .matches
        ? 'auto'
        : 'smooth';
      let expectedPosition: number;
      if (usesInternalScrollRef.current) {
        const scrollerRect = scroller.getBoundingClientRect();
        const targetTop =
          scroller.scrollTop +
          element.getBoundingClientRect().top -
          scrollerRect.top;
        expectedPosition = Math.min(
          Math.max(0, targetTop - OUTLINE_TARGET_TOP_OFFSET),
          Math.max(0, scroller.scrollHeight - scroller.clientHeight),
        );
        scroller.scrollTo({
          top: expectedPosition,
          behavior,
        });
      } else {
        expectedPosition = Math.min(
          Math.max(0, window.scrollY + element.getBoundingClientRect().top),
          Math.max(
            0,
            document.documentElement.scrollHeight - window.innerHeight,
          ),
        );
        element.scrollIntoView({ behavior, block: 'start' });
      }
      commitActiveKey(item.id);
      settleProgrammaticNavigation(element, expectedPosition);
    },
    [commitActiveKey, scrollContainerRef, settleProgrammaticNavigation],
  );

  const selectTarget = useCallback(
    (item: WordOutlineItem) => {
      navigationAbortRef.current?.abort();
      cancelNavigationSettle();
      const controller = new AbortController();
      navigationAbortRef.current = controller;
      programmaticNavigationRef.current = true;
      commitActiveKey(item.id);

      void (async () => {
        try {
          if (
            pageMode === 'windowed' &&
            pageSource &&
            blockPageIndex &&
            pageNavigationControllerRef?.current
          ) {
            const prioritizedPage = await pageSource.prioritizeBlock(
              item.targetBlockId,
              controller.signal,
            );
            const indexedPage = blockPageIndex.locate(
              item.targetBlockId,
            )?.pageIndex;
            const pageIndex = indexedPage ?? prioritizedPage;
            if (pageIndex >= 0) {
              const pageNavigation = pageNavigationControllerRef.current;
              if (!pageNavigation) return;
              pageNavigation.scrollToPage(pageIndex);
              const pageElement = await pageNavigation.ensurePageMounted(
                pageIndex,
                controller.signal,
              );
              const element = findOutlineElement(
                pageElement,
                item.targetBlockId,
              );
              if (element) {
                scrollToElement(element, item);
                return;
              }
            }
          }

          const scroller = scrollContainerRef.current;
          const direct = scroller
            ? findOutlineElement(scroller, item.targetBlockId)
            : undefined;
          if (direct) {
            scrollToElement(direct, item);
            return;
          }
          const cached = targetsRef.current.find(
            (candidate) => candidate.key === item.id,
          );
          if (cached) scrollToElement(cached.element, item);
          else programmaticNavigationRef.current = false;
        } catch (error) {
          if (!(error instanceof Error && error.name === 'AbortError')) {
            programmaticNavigationRef.current = false;
          }
        }
      })();
    },
    [
      blockPageIndex,
      cancelNavigationSettle,
      commitActiveKey,
      pageMode,
      pageNavigationControllerRef,
      pageSource,
      scrollContainerRef,
      scrollToElement,
    ],
  );

  return { activeKey, selectTarget };
}
