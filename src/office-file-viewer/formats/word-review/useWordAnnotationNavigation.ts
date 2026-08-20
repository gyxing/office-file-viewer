import type { MutableRefObject, RefObject } from 'react';
import { useCallback, useEffect, useRef } from 'react';
import type { OfficeAnnotation } from '../../services/annotations/types';
import type { WordPageSource } from '../../services/word/WordPageSource';
import { useOfficeAnnotationNavigation } from '../../shared/annotations';
import type { WordPageNavigationController } from '../word-pages/types';
import type { WordBlockPageIndex } from '../word-pages/WordBlockPageIndex';

/** 批注定位后在视口上下保留的最小空白。 */
const WORD_ANNOTATION_VIEWPORT_PADDING = 24;

/** 在当前页面中查找批注或修订范围，缺失时回退到所属正文块。 */
function findReviewTarget(root: ParentNode, reviewId: string, blockId: string) {
  const marker = Array.from(
    root.querySelectorAll<HTMLElement>('[data-office-annotation-id]'),
  ).find((element) => element.dataset.officeAnnotationId === reviewId);
  if (marker) return marker;
  const revision = Array.from(
    root.querySelectorAll<HTMLElement>('[data-office-word-revision]'),
  ).find((element) => element.dataset.officeWordRevision === reviewId);
  if (revision) return revision;
  return Array.from(
    root.querySelectorAll<HTMLElement>('[data-office-word-block-id]'),
  ).find((element) => element.dataset.officeWordBlockId === blockId);
}

/** 将批注范围平滑滚动到正文视口中部。 */
function scrollAnnotationIntoView(scroller: HTMLElement, element: HTMLElement) {
  const scrollerRect = scroller.getBoundingClientRect();
  const elementRect = element.getBoundingClientRect();
  const targetTop =
    scroller.scrollTop +
    elementRect.top -
    scrollerRect.top -
    Math.max(
      WORD_ANNOTATION_VIEWPORT_PADDING,
      (scroller.clientHeight - elementRect.height) / 2,
    );
  scroller.scrollTo({
    top: Math.min(
      Math.max(0, targetTop),
      Math.max(0, scroller.scrollHeight - scroller.clientHeight),
    ),
    behavior: window.matchMedia?.('(prefers-reduced-motion: reduce)').matches
      ? 'auto'
      : 'smooth',
  });
}

/** 注册普通分页和虚拟分页共用的 Word 批注定位。 */
export function useWordAnnotationNavigation({
  scrollContainerRef,
  pageMode,
  pageSource,
  blockPageIndex,
  pageNavigationControllerRef,
  documentSessionId,
}: {
  /** Word 正文滚动容器。 */
  scrollContainerRef: RefObject<HTMLElement>;
  /** 当前页面采用普通或窗口化渲染。 */
  pageMode: 'normal' | 'windowed';
  /** 按需提供页面内容的数据源。 */
  pageSource: WordPageSource<unknown>;
  /** 正文块到最新页面索引的查询结构。 */
  blockPageIndex: WordBlockPageIndex;
  /** 虚拟页面挂载与滚动控制器。 */
  pageNavigationControllerRef: MutableRefObject<
    WordPageNavigationController | undefined
  >;
  /** 当前文档解析会话标识。 */
  documentSessionId: string;
}) {
  const navigationAbortRef = useRef<AbortController>();
  const navigate = useCallback(
    async (annotation: OfficeAnnotation) => {
      if (annotation.target.kind !== 'word-range') return false;
      const scroller = scrollContainerRef.current;
      if (!scroller) return false;
      navigationAbortRef.current?.abort();
      const controller = new AbortController();
      navigationAbortRef.current = controller;
      try {
        let root: ParentNode = scroller;
        if (pageMode === 'windowed' && pageNavigationControllerRef.current) {
          const prioritizedPage = await pageSource.prioritizeBlock(
            annotation.target.blockId,
            controller.signal,
          );
          const pageIndex =
            blockPageIndex.locate(annotation.target.blockId)?.pageIndex ??
            annotation.target.pageIndex ??
            prioritizedPage;
          if (pageIndex >= 0) {
            const pageNavigation = pageNavigationControllerRef.current;
            pageNavigation.scrollToPage(pageIndex);
            root = await pageNavigation.ensurePageMounted(
              pageIndex,
              controller.signal,
            );
          }
        }
        const element = findReviewTarget(
          root,
          annotation.id,
          annotation.target.blockId,
        );
        if (!element) return false;
        scrollAnnotationIntoView(scroller, element);
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

  useOfficeAnnotationNavigation('word-range', navigate);
  useEffect(
    () => () => {
      navigationAbortRef.current?.abort();
      navigationAbortRef.current = undefined;
    },
    [documentSessionId],
  );
}
