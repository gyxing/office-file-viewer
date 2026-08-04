// office-file-viewer 模块的公共入口，业务侧通过这里使用 OfficeFileViewer 及相关类型。
export type { OfficeFileViewerLocale } from './locale';
export { OfficeFileViewer } from './OfficeFileViewer';
export type {
  OfficeFileViewerProps,
  OfficeFileViewerUri,
} from './OfficeFileViewer';
export { disposeDocDocument } from './services/doc/types';
export type { DocDocument, DocResources } from './services/doc/types';
export { createOfficeParseSession } from './services/parsing';
export type {
  OfficeParseOptions,
  OfficeParseSession,
  OfficeParseSessionStatus,
  OfficePreviewReadyInfo,
  ParseProgress,
  ParseStage,
  WorkerMode,
} from './services/parsing';
export { parsePpt } from './services/ppt';
export { disposePresentationDocument } from './services/presentation/dispose';
export type { PresentationDocument } from './services/presentation/types';
export { disposeParsedOfficeFile } from './services/preview';
export type { ParsedOfficeFile, PreviewKind } from './services/preview';
export { disposeSpreadsheetWorkbook } from './services/spreadsheet/types';
export type {
  SpreadsheetCell,
  SpreadsheetCellStyle,
  SpreadsheetResources,
  SpreadsheetSheet,
  SpreadsheetWarning,
  SpreadsheetWorkbook,
} from './services/spreadsheet/types';
