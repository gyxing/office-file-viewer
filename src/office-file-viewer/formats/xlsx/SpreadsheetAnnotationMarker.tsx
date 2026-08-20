import React from 'react';
import type { SpreadsheetAnnotation } from '../../services/spreadsheet/semantics/types';
import { useOfficeAnnotationRuntime } from '../../shared/annotations';

/** 单元格批注角标属性。 */
type SpreadsheetAnnotationMarkerProps = {
  /** 当前单元格关联的批注。 */
  annotation?: SpreadsheetAnnotation;
};

/** 渲染 Excel 风格的右上角批注标记，并打开共享审阅面板目标。 */
export function SpreadsheetAnnotationMarker({
  annotation,
}: SpreadsheetAnnotationMarkerProps) {
  const runtime = useOfficeAnnotationRuntime();
  if (!annotation || !runtime?.options.showComments) return null;
  return (
    <button
      type="button"
      className="office-file-spreadsheet-annotation-marker"
      aria-label={
        annotation.author
          ? `${annotation.author}：${annotation.text}`
          : annotation.text
      }
      data-office-spreadsheet-annotation={annotation.id}
      data-active={
        runtime.state.activeAnnotation?.id === annotation.id ? 'true' : 'false'
      }
      onClick={(event) => {
        event.stopPropagation();
        void runtime.actions.selectId(annotation.id);
      }}
    />
  );
}
