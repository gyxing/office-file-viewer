import type { RefObject } from 'react';
import { useCallback } from 'react';
import type { OfficeAnnotation } from '../../services/annotations/types';
import { useOfficeAnnotationNavigation } from '../../shared/annotations';

function escapeAttribute(value: string) {
  return value.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

/** 注册批注导航：先切换幻灯片，再聚焦批注标签或源绘制对象。 */
export function usePresentationAnnotationNavigation({
  viewerRef,
  onSelectSlide,
}: {
  /** 当前演示文稿根节点。 */
  viewerRef: RefObject<HTMLDivElement>;
  /** 切换到批注所属幻灯片。 */
  onSelectSlide(index: number): void;
}) {
  const navigate = useCallback(
    async (annotation: OfficeAnnotation) => {
      if (annotation.target.kind !== 'presentation-element') return false;
      const index = annotation.target.slideIndex;
      if (index === undefined) return false;
      onSelectSlide(index);
      for (let attempt = 0; attempt < 30; attempt += 1) {
        await new Promise<void>((resolve) => {
          requestAnimationFrame(() => resolve());
        });
        const root = viewerRef.current;
        if (!root) return false;
        const marker = root.querySelector<HTMLElement>(
          `.office-file-pptx-viewer__viewport [data-office-annotation-id="${escapeAttribute(
            annotation.id,
          )}"]`,
        );
        const element = annotation.target.elementId
          ? root.querySelector<HTMLElement>(
              `.office-file-pptx-viewer__viewport [data-office-presentation-element-id="${escapeAttribute(
                annotation.target.elementId,
              )}"]`,
            )
          : undefined;
        const target = marker ?? element;
        if (!target) continue;
        target.scrollIntoView({ block: 'center', inline: 'center' });
        marker?.focus({ preventScroll: true });
        return true;
      }
      return false;
    },
    [onSelectSlide, viewerRef],
  );
  useOfficeAnnotationNavigation('presentation-element', navigate);
}
