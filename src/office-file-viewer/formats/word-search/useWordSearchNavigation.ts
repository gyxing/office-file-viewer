import type { MutableRefObject, RefObject } from 'react';
import { useCallback, useEffect, useRef } from 'react';
import type { OfficeSearchTarget } from '../../services/search/types';
import type { WordPageSource } from '../../services/word/WordPageSource';
import { useOfficeSearchNavigatorRegistration } from '../search/OfficeSearchContext';
import type { WordPageNavigationController } from '../word-pages/types';
import type { WordBlockPageIndex } from '../word-pages/WordBlockPageIndex';

/** 查找结果滚动后保留的最小视口边距。 */
const WORD_SEARCH_VIEWPORT_PADDING = 24;

function findWordBlockElement(root: ParentNode, blockId: string) {
  return Array.from(
    root.querySelectorAll<HTMLElement>('[data-office-word-block-id]'),
  ).find((element) => element.dataset.officeWordBlockId === blockId);
}

function scrollWordBlockIntoView(scroller: HTMLElement, element: HTMLElement) {
  const scrollerRect = scroller.getBoundingClientRect();
  const elementRect = element.getBoundingClientRect();
  const targetTop =
    scroller.scrollTop +
    elementRect.top -
    scrollerRect.top -
    Math.max(
      WORD_SEARCH_VIEWPORT_PADDING,
      (scroller.clientHeight - elementRect.height) / 2,
    );
  const maximumTop = Math.max(0, scroller.scrollHeight - scroller.clientHeight);
  scroller.scrollTo({
    top: Math.min(Math.max(0, targetTop), maximumTop),
    behavior: window.matchMedia?.('(prefers-reduced-motion: reduce)').matches
      ? 'auto'
      : 'smooth',
  });
}

/** 注册 Word 查找结果定位，并复用大文件虚拟分页控制器按需挂载目标页。 */
export function useWordSearchNavigation({
  scrollContainerRef,
  pageMode,
  pageSource,
  blockPageIndex,
  pageNavigationControllerRef,
  documentSessionId,
}: {
  /** 正文滚动容器。 */
  scrollContainerRef: RefObject<HTMLElement>;
  /** 当前使用的页面加载模式。 */
  pageMode: 'normal' | 'windowed';
  /** 按需提供页面内容的数据源。 */
  pageSource: WordPageSource<unknown>;
  /** 正文块到页面索引的查询结构。 */
  blockPageIndex: WordBlockPageIndex;
  /** 虚拟页挂载和滚动控制器。 */
  pageNavigationControllerRef: MutableRefObject<
    WordPageNavigationController | undefined
  >;
  /** 当前文档解析会话的标识。 */
  documentSessionId: string;
}) {
  const navigationAbortRef = useRef<AbortController>();
  const navigate = useCallback(
    async (target: OfficeSearchTarget) => {
      if (target.kind !== 'word') return false;
      const scroller = scrollContainerRef.current;
      if (!scroller) return false;

      navigationAbortRef.current?.abort();
      const controller = new AbortController();
      navigationAbortRef.current = controller;
      try {
        let searchRoot: ParentNode = scroller;
        if (pageMode === 'windowed' && pageNavigationControllerRef.current) {
          const prioritizedPage = await pageSource.prioritizeBlock(
            target.blockId,
            controller.signal,
          );
          const pageIndex =
            blockPageIndex.locate(target.blockId)?.pageIndex ??
            target.pageIndex ??
            prioritizedPage;
          if (pageIndex >= 0) {
            const pageNavigation = pageNavigationControllerRef.current;
            pageNavigation.scrollToPage(pageIndex);
            searchRoot = await pageNavigation.ensurePageMounted(
              pageIndex,
              controller.signal,
            );
          }
        }

        const element = findWordBlockElement(searchRoot, target.blockId);
        if (!element) return false;
        scrollWordBlockIntoView(scroller, element);
        return true;
      } catch (error) {
        if (error instanceof Error && error.name === 'AbortError') return false;
        return false;
      }
    },
    [
      blockPageIndex,
      pageMode,
      pageNavigationControllerRef,
      pageSource,
      scrollContainerRef,
    ],
  );

  useOfficeSearchNavigatorRegistration('word', navigate);

  useEffect(
    () => () => {
      navigationAbortRef.current?.abort();
      navigationAbortRef.current = undefined;
    },
    [documentSessionId],
  );
}
