import type { MutableRefObject, RefObject } from 'react';
import { useCallback, useEffect, useRef } from 'react';
import type { WordBookmarkTarget } from '../../services/word/types';
import type { WordPageSource } from '../../services/word/WordPageSource';
import type { OfficeInternalHyperlinkTarget } from '../../shared/hyperlink';
import { useOfficeHyperlinkContext } from '../../shared/hyperlink';
import type { WordPageNavigationController } from '../word-pages/types';
import type { WordBlockPageIndex } from '../word-pages/WordBlockPageIndex';

/** Word 内部书签导航使用的页面与滚动能力。 */
type UseWordTargetNavigationOptions = {
  /** 按源名称索引的书签目标。 */
  bookmarks?: Record<string, WordBookmarkTarget>;
  /** 正文滚动容器。 */
  scrollContainerRef: RefObject<HTMLElement>;
  /** 当前使用的页面加载模式。 */
  pageMode: 'normal' | 'windowed';
  /** 按需提供页面内容的数据源。 */
  pageSource?: WordPageSource<unknown>;
  /** 正文块到页面索引的查询结构。 */
  blockPageIndex?: WordBlockPageIndex;
  /** 虚拟页挂载和滚动控制器。 */
  pageNavigationControllerRef?: MutableRefObject<
    WordPageNavigationController | undefined
  >;
  /** 当前文档解析会话的标识。 */
  documentSessionId: string;
};

/** 链接定位后保留顶部间距，避免目标紧贴工具栏。 */
const WORD_HYPERLINK_TOP_OFFSET = 20;

function findBookmarkElement(root: ParentNode, name: string) {
  return Array.from(
    root.querySelectorAll<HTMLElement>('[data-office-word-bookmark]'),
  ).find((element) => element.dataset.officeWordBookmark === name);
}

function scrollToBookmark(scroller: HTMLElement, element: HTMLElement) {
  const behavior = window.matchMedia?.('(prefers-reduced-motion: reduce)')
    .matches
    ? 'auto'
    : 'smooth';
  const scrollerRect = scroller.getBoundingClientRect();
  const targetTop =
    scroller.scrollTop + element.getBoundingClientRect().top - scrollerRect.top;
  scroller.scrollTo({
    top: Math.min(
      Math.max(0, targetTop - WORD_HYPERLINK_TOP_OFFSET),
      Math.max(0, scroller.scrollHeight - scroller.clientHeight),
    ),
    behavior,
  });
  element.tabIndex = -1;
  element.focus({ preventScroll: true });
}

/** 注册 Word 内部书签导航，并兼容普通页面与虚拟页面。 */
export function useWordTargetNavigation({
  bookmarks,
  scrollContainerRef,
  pageMode,
  pageSource,
  blockPageIndex,
  pageNavigationControllerRef,
  documentSessionId,
}: UseWordTargetNavigationOptions) {
  const hyperlinkContext = useOfficeHyperlinkContext();
  const navigationAbortRef = useRef<AbortController>();

  const navigate = useCallback(
    async (target: OfficeInternalHyperlinkTarget) => {
      if (target.family !== 'word') return false;
      const bookmark = bookmarks?.[target.bookmark];
      const scroller = scrollContainerRef.current;
      if (!bookmark || !scroller) return false;

      navigationAbortRef.current?.abort();
      const controller = new AbortController();
      navigationAbortRef.current = controller;
      try {
        if (
          pageMode === 'windowed' &&
          pageSource &&
          blockPageIndex &&
          pageNavigationControllerRef?.current
        ) {
          const prioritizedPage = await pageSource.prioritizeBlock(
            bookmark.targetBlockId,
            controller.signal,
          );
          const pageIndex =
            blockPageIndex.locate(bookmark.targetBlockId)?.pageIndex ??
            prioritizedPage;
          if (pageIndex >= 0) {
            const pageNavigation = pageNavigationControllerRef.current;
            pageNavigation.scrollToPage(pageIndex);
            const pageElement = await pageNavigation.ensurePageMounted(
              pageIndex,
              controller.signal,
            );
            const element = findBookmarkElement(pageElement, bookmark.name);
            if (element) {
              scrollToBookmark(scroller, element);
              return true;
            }
          }
        }

        const element = findBookmarkElement(scroller, bookmark.name);
        if (!element) return false;
        scrollToBookmark(scroller, element);
        return true;
      } catch (error) {
        if (error instanceof Error && error.name === 'AbortError') return false;
        return false;
      }
    },
    [
      blockPageIndex,
      bookmarks,
      pageMode,
      pageNavigationControllerRef,
      pageSource,
      scrollContainerRef,
    ],
  );

  useEffect(() => {
    if (!hyperlinkContext) return undefined;
    return hyperlinkContext.registerNavigator('word', navigate);
  }, [hyperlinkContext, navigate]);

  useEffect(
    () => () => {
      navigationAbortRef.current?.abort();
      navigationAbortRef.current = undefined;
    },
    [documentSessionId],
  );
}
