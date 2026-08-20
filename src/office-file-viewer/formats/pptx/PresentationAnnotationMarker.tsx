import type { CSSProperties } from 'react';
import React from 'react';
import type { PresentationAnnotation } from '../../services/presentation/types';
import { useOfficeAnnotationRuntime } from '../../shared/annotations';

/** 幻灯片批注标签属性。 */
type PresentationAnnotationMarkerProps = {
  /** 当前幻灯片中的批注。 */
  annotation: PresentationAnnotation;
  /** 缺少坐标时用于错开多个页级批注。 */
  ordinal: number;
  /** 主视口可点击，缩略图只显示非纯颜色标记。 */
  interactive: boolean;
};

/** 渲染幻灯片批注标签，并与共享审阅面板保持同一活动项。 */
export function PresentationAnnotationMarker({
  annotation,
  ordinal,
  interactive,
}: PresentationAnnotationMarkerProps) {
  const runtime = useOfficeAnnotationRuntime();
  if (!runtime?.options.showComments) return null;
  const active = runtime.state.activeAnnotation?.id === annotation.id;
  const style: CSSProperties = {
    left: annotation.x ?? undefined,
    top: annotation.y ?? 12 + ordinal * 24,
    right: annotation.x === undefined ? 12 : undefined,
  };
  const className = [
    'office-file-pptx-annotation-marker',
    active ? 'office-file-pptx-annotation-marker--active' : undefined,
    interactive ? undefined : 'office-file-pptx-annotation-marker--thumbnail',
  ]
    .filter(Boolean)
    .join(' ');
  const label = annotation.author
    ? `${annotation.author}：${annotation.text}`
    : annotation.text;
  if (!interactive) {
    return <span className={className} style={style} aria-hidden="true" />;
  }
  return (
    <button
      type="button"
      className={className}
      style={style}
      aria-label={label}
      data-office-annotation-id={annotation.id}
      data-active={active ? 'true' : 'false'}
      onClick={(event) => {
        event.stopPropagation();
        void runtime.actions.selectId(annotation.id);
      }}
    >
      {ordinal + 1}
    </button>
  );
}
