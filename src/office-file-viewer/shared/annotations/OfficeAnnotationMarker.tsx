import type { ReactNode } from 'react';
import React from 'react';
import { useOfficeAnnotationRuntime } from './OfficeAnnotationProvider';

/** 正文批注范围标记属性。 */
type OfficeAnnotationMarkerProps = {
  /** 当前范围对应的批注标识。 */
  annotationId: string;
  /** 不改变原有文本排版的行内内容。 */
  children: ReactNode;
  /** 源格式已经标记为解决时降低视觉强调。 */
  resolved?: boolean;
  /** 当前 Word 投影是否显示批注标记。 */
  visible?: boolean;
};

/** 渲染可选择、可导航且不会改变行高的批注范围。 */
export function OfficeAnnotationMarker({
  annotationId,
  children,
  resolved = false,
  visible = true,
}: OfficeAnnotationMarkerProps) {
  const runtime = useOfficeAnnotationRuntime();
  if (!visible || !runtime?.options.showComments) return <>{children}</>;
  const active = runtime.state.activeAnnotation?.id === annotationId;

  return (
    <mark
      className={[
        'office-file-annotation-marker',
        active ? 'office-file-annotation-marker--active' : undefined,
        resolved ? 'office-file-annotation-marker--resolved' : undefined,
      ]
        .filter(Boolean)
        .join(' ')}
      data-office-annotation-id={annotationId}
      data-active={active ? 'true' : 'false'}
      onClick={() => void runtime.actions.selectId(annotationId)}
    >
      {children}
    </mark>
  );
}
