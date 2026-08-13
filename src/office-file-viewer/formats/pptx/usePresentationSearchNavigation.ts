import type { RefObject } from 'react';
import { useCallback, useEffect, useRef } from 'react';
import type { PresentationSourceSnapshot } from '../../services/presentation/PresentationSource';
import type { OfficeSearchTarget } from '../../services/search/types';
import { useOfficeSearchNavigatorRegistration } from '../search/OfficeSearchContext';

/** 切换幻灯片后等待目标元素挂载的最大帧数。 */
const PRESENTATION_SEARCH_NAVIGATION_MAX_FRAMES = 90;

function findPresentationElement(root: ParentNode, elementId: string) {
  return Array.from(
    root.querySelectorAll<HTMLElement>('[data-office-presentation-element-id]'),
  ).find(
    (element) => element.dataset.officePresentationElementId === elementId,
  );
}

function waitForPresentationElement(
  rootRef: RefObject<HTMLElement>,
  elementId: string,
  isCurrent: () => boolean,
) {
  return new Promise<HTMLElement | undefined>((resolve) => {
    let frame = 0;
    const inspect = () => {
      if (!isCurrent()) {
        resolve(undefined);
        return;
      }
      const viewport = rootRef.current?.querySelector<HTMLElement>(
        '.office-file-pptx-viewer__viewport',
      );
      const element = viewport
        ? findPresentationElement(viewport, elementId)
        : undefined;
      if (element) {
        resolve(element);
        return;
      }
      frame += 1;
      if (frame >= PRESENTATION_SEARCH_NAVIGATION_MAX_FRAMES) {
        resolve(undefined);
        return;
      }
      requestAnimationFrame(inspect);
    };
    inspect();
  });
}

function scrollPresentationElementIntoView(
  root: HTMLElement,
  element: HTMLElement,
) {
  const viewport = root.querySelector<HTMLElement>(
    '.office-file-pptx-viewer__viewport',
  );
  if (!viewport) return;
  const viewportRect = viewport.getBoundingClientRect();
  const elementRect = element.getBoundingClientRect();
  const maximumLeft = Math.max(0, viewport.scrollWidth - viewport.clientWidth);
  const maximumTop = Math.max(0, viewport.scrollHeight - viewport.clientHeight);
  viewport.scrollTo({
    left: Math.min(
      maximumLeft,
      Math.max(
        0,
        viewport.scrollLeft +
          elementRect.left -
          viewportRect.left -
          (viewport.clientWidth - elementRect.width) / 2,
      ),
    ),
    top: Math.min(
      maximumTop,
      Math.max(
        0,
        viewport.scrollTop +
          elementRect.top -
          viewportRect.top -
          (viewport.clientHeight - elementRect.height) / 2,
      ),
    ),
    behavior: window.matchMedia?.('(prefers-reduced-motion: reduce)').matches
      ? 'auto'
      : 'smooth',
  });
}

/** 注册演示文稿查找结果导航，通过现有受控选页状态切换到目标幻灯片。 */
export function usePresentationSearchNavigation({
  snapshot,
  onSelectSlide,
  viewerRef,
}: {
  /** 当前文稿的轻量幻灯片快照。 */
  snapshot: PresentationSourceSnapshot;
  /** 请求切换当前幻灯片。 */
  onSelectSlide: (index: number) => void;
  /** 演示文稿 Viewer 根元素。 */
  viewerRef: RefObject<HTMLElement>;
}) {
  const navigationGenerationRef = useRef(0);
  const navigate = useCallback(
    async (target: OfficeSearchTarget) => {
      if (
        target.kind !== 'presentation' ||
        target.slideIndex < 0 ||
        target.slideIndex >= snapshot.slideCount
      ) {
        return false;
      }
      const generation = ++navigationGenerationRef.current;
      onSelectSlide(target.slideIndex);
      const element = await waitForPresentationElement(
        viewerRef,
        target.elementId,
        () => generation === navigationGenerationRef.current,
      );
      const root = viewerRef.current;
      if (!root || !element) return false;
      scrollPresentationElementIntoView(root, element);
      return true;
    },
    [onSelectSlide, snapshot.slideCount, viewerRef],
  );

  useOfficeSearchNavigatorRegistration('presentation', navigate);

  useEffect(
    () => () => {
      navigationGenerationRef.current += 1;
    },
    [snapshot.revision],
  );
}
