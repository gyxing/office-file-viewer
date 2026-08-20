// office-file-viewer 模块的公共入口，业务侧通过这里使用 OfficeFileViewer 及相关类型。
import type { PresentationDocument } from './services/presentation/types';

export type { OfficeFileViewerLocale } from './locale';
export { OfficeFileViewer } from './OfficeFileViewer';
export type {
  OfficeFileViewerFontOptions,
  OfficeFileViewerPresentationMediaOptions,
  OfficeFileViewerPresentationTransitions,
  OfficeFileViewerProps,
  OfficeFileViewerSearchOptions,
  OfficeFileViewerUri,
  OfficeFileViewerUriLoader,
} from './OfficeFileViewer';
export type {
  OfficeAnnotation,
  OfficeAnnotationTarget,
  OfficeFileViewerReviewOptions,
  OfficePresentationAnnotationTarget,
  OfficeSpreadsheetAnnotationTarget,
  OfficeWordAnnotationTarget,
  WordRevisionMode,
} from './services/annotations';
export { disposeDocDocument } from './services/doc/types';
export type { DocDocument, DocResources } from './services/doc/types';
export {
  createOfficeParseSession,
  OfficeResourceLimitError,
} from './services/parsing';
export type {
  OfficeParseOptions,
  OfficeParseResourcePolicy,
  OfficeParseSession,
  OfficeParseSessionStatus,
  OfficePreviewReadyInfo,
  OfficeResourceLimitCode,
  ParseProgress,
  ParseStage,
  WorkerMode,
} from './services/parsing';
export type {
  OfficeHyperlink,
  OfficeHyperlinkActivateEvent,
  OfficeHyperlinkSourceType,
  OfficeInternalHyperlinkTarget,
} from './shared/hyperlink';
export type {
  OfficeFileViewerImagePreviewConfig,
  OfficeFileViewerImagePreviewOptions,
} from './shared/image-preview';
/** 按需加载 PPT 二进制解析器，避免仅使用预览组件时进入主包。 */
export async function parsePpt(file: File): Promise<PresentationDocument> {
  const { parsePpt: parsePptFile } = await import('./services/ppt/parsePpt');
  return parsePptFile(file);
}

export { disposePresentationDocument } from './services/presentation/dispose';
export type {
  PresentationMediaKind,
  PresentationMediaSource,
} from './services/presentation/mediaTypes';
export type {
  PresentationTransition,
  PresentationTransitionDirection,
  PresentationTransitionType,
} from './services/presentation/transitionTypes';
export type {
  PresentationAnnotation,
  PresentationDocument,
} from './services/presentation/types';
export { disposeParsedOfficeFile } from './services/preview';
export type { ParsedOfficeFile, PreviewKind } from './services/preview';
export type {
  OfficeFileViewerWarning,
  OfficeFileViewerWarningSource,
  OfficeFontFallbackWarning,
} from './services/previewWarnings';
export { disposeSpreadsheetWorkbook } from './services/spreadsheet/types';
export type {
  SpreadsheetCell,
  SpreadsheetCellStyle,
  SpreadsheetResources,
  SpreadsheetSheet,
  SpreadsheetWarning,
  SpreadsheetWorkbook,
} from './services/spreadsheet/types';
export type { SpreadsheetViewMode } from './services/spreadsheet/viewMode';
export type {
  OfficeFileViewerViewState,
  OfficeFileViewerViewStateChange,
} from './shell/viewState';
