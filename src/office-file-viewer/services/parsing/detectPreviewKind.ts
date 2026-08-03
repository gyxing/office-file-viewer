import {
  getPreviewFamily,
  tryDetectPreviewKind,
  type PreviewKind,
} from './formatDefinitions';

export {
  getOfficeFormatMetadata,
  getPreviewFamily,
  OFFICE_FORMAT_METADATA,
  SUPPORTED_OFFICE_EXTENSIONS,
  tryDetectPreviewKind,
} from './formatDefinitions';
export type {
  OfficeFormatMetadata,
  PreviewFamily,
  PreviewKind,
} from './formatDefinitions';

/** 判断文件名是否属于支持的 Office 格式。 */
export function isSupportedOfficeFileName(fileName: string): boolean {
  return tryDetectPreviewKind(fileName) !== undefined;
}

/** 判断当前格式是否复用电子表格预览链路。 */
export function isSpreadsheetPreviewKind(
  kind: PreviewKind,
): kind is 'xlsx' | 'xls' {
  return getPreviewFamily(kind) === 'spreadsheet';
}

/** 判断当前格式是否复用统一文字文档预览链路。 */
export function isWordPreviewKind(kind: PreviewKind): kind is 'docx' | 'doc' {
  return getPreviewFamily(kind) === 'word';
}

/** 判断当前格式是否复用统一演示文稿渲染链路。 */
export function isPresentationPreviewKind(
  kind: PreviewKind,
): kind is 'pptx' | 'ppt' {
  return getPreviewFamily(kind) === 'presentation';
}

/** 根据文件名推断 Office 预览格式。 */
export function detectPreviewKind(fileName: string): PreviewKind {
  // 保留历史行为：无可识别扩展名时按 PPTX 尝试解析。
  return tryDetectPreviewKind(fileName) ?? 'pptx';
}
