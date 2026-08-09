import type { OfficeFileViewerPreviewState } from './parsing/internalTypes';
import type { PreviewKind } from './preview';

/** 区分解析器警告与保留部分预览的运行时降级。 */
export type OfficeFileViewerWarningSource =
  | 'parser'
  | 'partial-preview'
  | 'hyperlink';

/** OfficeFileViewer 向宿主报告的非致命问题。 */
export type OfficeFileViewerWarning = {
  /** 供宿主稳定识别警告类型的代码。 */
  code: string;
  /** 面向日志或界面展示的警告说明。 */
  message: string;
  /** 产生警告的文件格式。 */
  previewKind: PreviewKind;
  /** 警告来自解析器、部分预览降级或链接运行时。 */
  source: OfficeFileViewerWarningSource;
};

/** 读取物化模型或按需数据源已经公开的解析警告。 */
export function collectOfficePreviewWarnings(
  preview: OfficeFileViewerPreviewState,
): OfficeFileViewerWarning[] {
  let warnings: ReadonlyArray<{ code: string; message: string }> = [];
  if (preview.mode === 'source') {
    if (preview.previewKind === 'doc') {
      warnings = preview.summary.warnings.map((message) => ({
        code: 'DOC_PARSE_WARNING',
        message,
      }));
    } else if (
      preview.previewKind === 'ppt' ||
      preview.previewKind === 'pptx'
    ) {
      warnings = preview.summary.warnings ?? [];
    }
  } else if (preview.model.kind === 'doc') {
    warnings = preview.model.document.warnings.map((message) => ({
      code: 'DOC_PARSE_WARNING',
      message,
    }));
  } else if (preview.model.kind === 'ppt' || preview.model.kind === 'pptx') {
    warnings = preview.model.document.warnings ?? [];
  } else if (preview.model.kind === 'xls' || preview.model.kind === 'xlsx') {
    warnings = preview.model.workbook.warnings ?? [];
  }

  const seen = new Set<string>();
  return warnings.flatMap((warning) => {
    const identity = `${warning.code}\u0000${warning.message}`;
    if (seen.has(identity)) return [];
    seen.add(identity);
    return [
      {
        code: warning.code,
        message: warning.message,
        previewKind: preview.previewKind,
        source: 'parser' as const,
      },
    ];
  });
}
