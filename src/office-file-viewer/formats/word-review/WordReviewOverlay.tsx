import type { CSSProperties } from 'react';
import React, {
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import { createPortal } from 'react-dom';
import { useOfficeFileViewerMessages } from '../../locale';
import type { OfficeAnnotation } from '../../services/annotations/types';
import type { WordRevisionKind } from '../../services/word/review/types';
import { useOfficeAnnotationRuntime } from '../../shared/annotations';
import {
  collectVisibleAnnotationIds,
  collectVisibleCommentAnchors,
  collectVisibleRevisionAnchors,
  intersectsBoundary,
  type VisibleScrollerBoundary,
} from './wordMarkupAnchors';
import { WordMarkupCallout } from './WordMarkupCallout';
import type {
  WordMarkupCalloutBoundary,
  WordMarkupCalloutLayout,
} from './wordMarkupCalloutLayout';
import {
  getWordMarkupRailExtent,
  layoutWordMarkupCallouts,
} from './wordMarkupCalloutLayout';
import { useWordRevisionMode } from './WordRevisionText';

/** 将动态坐标约束到页面安全区域。 */
function clamp(value: number, minimum: number, maximum: number) {
  return Math.min(Math.max(value, minimum), Math.max(minimum, maximum));
}

/** 避免滚动测量结果未变化时重复刷新标记数据。 */
function areStringArraysEqual(
  current: readonly string[],
  next: readonly string[],
) {
  return (
    current.length === next.length &&
    current.every((value, index) => value === next[index])
  );
}

/** 避免相同页侧布局在滚动过程中重复触发渲染。 */
function areMarkupCalloutsEqual(
  current: readonly WordMarkupCalloutLayout[],
  next: readonly WordMarkupCalloutLayout[],
) {
  return (
    current.length === next.length &&
    current.every((callout, index) => {
      const nextCallout = next[index];
      return (
        callout.key === nextCallout.key &&
        callout.author === nextCallout.author &&
        callout.excerpt === nextCallout.excerpt &&
        callout.resolved === nextCallout.resolved &&
        callout.calloutLeft === nextCallout.calloutLeft &&
        callout.calloutTop === nextCallout.calloutTop &&
        callout.anchorX === nextCallout.anchorX &&
        callout.anchorY === nextCallout.anchorY &&
        callout.connectorY === nextCallout.connectorY
      );
    })
  );
}

/** 在 Word 页面右侧统一展示批注和修订标记。 */
export function WordReviewOverlay() {
  const messages = useOfficeFileViewerMessages();
  const runtime = useOfficeAnnotationRuntime();
  const revisionMode = useWordRevisionMode();
  const overlayRef = useRef<HTMLDivElement>(null);
  const [calloutHost, setCalloutHost] = useState<HTMLElement>();
  const [annotationIds, setAnnotationIds] = useState<readonly string[]>([]);
  const [annotations, setAnnotations] = useState<readonly OfficeAnnotation[]>(
    [],
  );
  const [hoveredMarkupKey, setHoveredMarkupKey] = useState<string>();
  const [callouts, setCallouts] = useState<readonly WordMarkupCalloutLayout[]>(
    [],
  );
  const [railExtent, setRailExtent] = useState(0);
  const activeRevision = runtime?.state.activeRevision;
  const activeAnnotation = runtime?.state.activeAnnotation;
  const loadAnnotationsByIds = runtime?.actions.loadAnnotationsByIds;
  const annotationById = useMemo(
    () => new Map(annotations.map((annotation) => [annotation.id, annotation])),
    [annotations],
  );

  useEffect(() => {
    if (
      revisionMode !== 'markup' ||
      !annotationIds.length ||
      !loadAnnotationsByIds
    ) {
      setAnnotations((current) => (current.length ? [] : current));
      return undefined;
    }
    let active = true;
    void loadAnnotationsByIds(annotationIds).then((loaded) => {
      if (active) setAnnotations(loaded);
    });
    return () => {
      active = false;
    };
  }, [annotationIds, loadAnnotationsByIds, revisionMode]);

  useEffect(() => {
    const article = overlayRef.current?.closest('article');
    const measurementArticle = article?.closest(
      '.office-file-docx-viewer__measure',
    );
    if (!article || measurementArticle || revisionMode !== 'markup') {
      setHoveredMarkupKey(undefined);
      return undefined;
    }
    const findMarkupKey = (target: EventTarget | null) => {
      if (!(target instanceof Element)) return undefined;
      const annotation = target.closest<HTMLElement>(
        '[data-office-annotation-id]',
      );
      if (annotation?.dataset.officeAnnotationId) {
        return `comment:${annotation.dataset.officeAnnotationId}`;
      }
      const revision = target.closest<HTMLElement>(
        '[data-office-word-revision]',
      );
      return revision?.dataset.officeWordRevision
        ? `revision:${revision.dataset.officeWordRevision}`
        : undefined;
    };
    const handlePointerOver = (event: Event) => {
      setHoveredMarkupKey(findMarkupKey(event.target));
    };
    const handlePointerOut = (event: MouseEvent) => {
      const previous = findMarkupKey(event.target);
      const next = findMarkupKey(event.relatedTarget);
      if (previous !== next) setHoveredMarkupKey(next);
    };
    article.addEventListener('mouseover', handlePointerOver);
    article.addEventListener('mouseout', handlePointerOut);
    return () => {
      article.removeEventListener('mouseover', handlePointerOver);
      article.removeEventListener('mouseout', handlePointerOut);
    };
  }, [revisionMode]);

  useLayoutEffect(() => {
    const overlay = overlayRef.current;
    const article = overlay?.closest('article');
    if (!overlay || !article) return undefined;
    const scroller = article.closest<HTMLElement>(
      '.office-file-docx-viewer__scroller, .office-file-doc-viewer__scroller',
    );
    const pageFrame = article.closest<HTMLElement>(
      '.office-file-docx-page-frame, .office-file-doc-page-frame',
    );
    setCalloutHost((current) =>
      current === scroller ? current : scroller ?? undefined,
    );
    let animationFrame = 0;
    const measure = () => {
      const isMeasurementArticle = Boolean(
        article.closest('.office-file-docx-viewer__measure'),
      );
      if (
        revisionMode !== 'markup' ||
        isMeasurementArticle ||
        window.getComputedStyle(article).visibility === 'hidden'
      ) {
        setAnnotationIds((current) => (current.length ? [] : current));
        setCallouts((current) => (current.length ? [] : current));
        setRailExtent(0);
        return;
      }
      const articleRect = article.getBoundingClientRect();
      const scrollerRect = scroller?.getBoundingClientRect();
      if (!scroller || !scrollerRect) return;
      const scrollerClientLeft = scrollerRect.left + scroller.clientLeft;
      const scrollerClientTop = scrollerRect.top + scroller.clientTop;
      const visibleBoundary: VisibleScrollerBoundary = {
        left: Math.max(0, scrollerClientLeft),
        top: Math.max(0, scrollerClientTop),
        right: Math.min(
          window.innerWidth,
          scrollerClientLeft + scroller.clientWidth,
        ),
        bottom: Math.min(
          window.innerHeight,
          scrollerClientTop + scroller.clientHeight,
        ),
      };
      if (!intersectsBoundary(articleRect, visibleBoundary)) {
        setAnnotationIds((current) => (current.length ? [] : current));
        setCallouts((current) => (current.length ? [] : current));
        setRailExtent(0);
        return;
      }
      const contentOriginLeft = scrollerClientLeft - scroller.scrollLeft;
      const contentOriginTop = scrollerClientTop - scroller.scrollTop;
      const visibleAnnotationIds = collectVisibleAnnotationIds(
        article,
        visibleBoundary,
      );
      setAnnotationIds((current) =>
        areStringArraysEqual(current, visibleAnnotationIds)
          ? current
          : visibleAnnotationIds,
      );
      const anchors = [
        ...collectVisibleCommentAnchors({
          article,
          boundary: visibleBoundary,
          contentOriginLeft,
          contentOriginTop,
          annotationById,
        }),
        ...collectVisibleRevisionAnchors({
          article,
          boundary: visibleBoundary,
          contentOriginLeft,
          contentOriginTop,
          activeRevision,
        }),
      ];
      const boundary: WordMarkupCalloutBoundary = {
        top: clamp(
          visibleBoundary.top - contentOriginTop,
          articleRect.top - contentOriginTop,
          articleRect.bottom - contentOriginTop,
        ),
        bottom: clamp(
          visibleBoundary.bottom - contentOriginTop,
          articleRect.top - contentOriginTop,
          articleRect.bottom - contentOriginTop,
        ),
        pageRight: articleRect.right - contentOriginLeft,
      };
      const nextCallouts = layoutWordMarkupCallouts(anchors, boundary);
      setCallouts((current) =>
        areMarkupCalloutsEqual(current, nextCallouts) ? current : nextCallouts,
      );
      setRailExtent(getWordMarkupRailExtent(boundary.pageRight, nextCallouts));
    };
    const scheduleMeasure = () => {
      if (animationFrame) return;
      animationFrame = window.requestAnimationFrame(() => {
        animationFrame = 0;
        measure();
      });
    };
    measure();
    const resizeObserver =
      typeof ResizeObserver === 'undefined'
        ? undefined
        : new ResizeObserver(scheduleMeasure);
    const mutationObserver =
      typeof MutationObserver === 'undefined'
        ? undefined
        : new MutationObserver(scheduleMeasure);
    resizeObserver?.observe(article);
    if (pageFrame) resizeObserver?.observe(pageFrame);
    if (scroller) {
      resizeObserver?.observe(scroller);
      mutationObserver?.observe(scroller, { childList: true });
      scroller.addEventListener('scroll', scheduleMeasure, { passive: true });
    }
    window.addEventListener('resize', scheduleMeasure, { passive: true });
    window.addEventListener('scroll', scheduleMeasure, {
      capture: true,
      passive: true,
    });
    return () => {
      if (animationFrame) window.cancelAnimationFrame(animationFrame);
      resizeObserver?.disconnect();
      mutationObserver?.disconnect();
      scroller?.removeEventListener('scroll', scheduleMeasure);
      window.removeEventListener('resize', scheduleMeasure);
      window.removeEventListener('scroll', scheduleMeasure, true);
    };
  }, [activeRevision, annotationById, revisionMode]);

  const getRevisionKindLabel = (kind: WordRevisionKind | undefined) => {
    switch (kind) {
      case 'insert':
        return messages.review.revisionInsert;
      case 'delete':
        return messages.review.revisionDelete;
      case 'move-from':
        return messages.review.revisionMoveFrom;
      case 'move-to':
        return messages.review.revisionMoveTo;
      case 'format':
        return messages.review.revisionFormat;
      default:
        return undefined;
    }
  };

  return (
    <div
      ref={overlayRef}
      className="office-file-word-review-overlay"
      data-markup-callout-count={callouts.length}
    >
      {calloutHost && revisionMode === 'markup'
        ? createPortal(
            <div className="office-file-word-markup-portal">
              <span
                className="office-file-word-markup-portal__extent"
                style={{ left: railExtent } as CSSProperties}
                aria-hidden="true"
              />
              {callouts.map((callout) => {
                const emphasized =
                  hoveredMarkupKey === callout.key ||
                  (callout.type === 'comment'
                    ? activeAnnotation?.id === callout.id
                    : activeRevision?.id === callout.id);
                return (
                  <WordMarkupCallout
                    key={callout.key}
                    callout={callout}
                    revisionKindLabel={getRevisionKindLabel(
                      callout.revisionKind,
                    )}
                    unknownAuthorLabel={messages.review.unknownAuthor}
                    emptyContentLabel={
                      callout.type === 'comment'
                        ? messages.review.emptyComment
                        : messages.review.emptyRevision
                    }
                    emphasized={emphasized}
                    onSelect={() => {
                      if (callout.type === 'comment') {
                        void runtime?.actions.selectId(callout.id);
                        return;
                      }
                      if (activeRevision?.id === callout.id) {
                        runtime?.actions.clearActiveRevision();
                      } else {
                        void runtime?.actions.activateRevisionId(callout.id);
                      }
                    }}
                  />
                );
              })}
            </div>,
            calloutHost,
          )
        : null}
    </div>
  );
}
