/** Office Core 公共入口，只导出与界面无关的稳定契约。 */
export type {
  OfficeCommand,
  OfficeCommandDefinition,
  OfficeEditorPluginAdapter,
  OfficeHistory,
  OfficeSelection,
  OfficeTransaction,
} from './editing';
export {
  createOfficeDocumentRuntime,
  createOfficeDocumentSnapshot,
  createOfficeSourceRuntime,
  createOfficeSourceSnapshot,
  getOfficeCapabilities,
  getOfficeDocumentModel,
} from './model';
export type {
  OfficeDocumentSnapshotInput,
  OfficeSourceSnapshotInput,
} from './model';
export { parseOfficeFile } from './parsing';
export type {
  OfficeFormatPlugin,
  OfficeParser,
  OfficeParserInput,
  OfficeParserPlugin,
  OfficePluginResolver,
  OfficeSourceFactory,
  OfficeWorkerParser,
  OfficeWorkerParserInput,
} from './parsingContracts';
export { createOfficeDocumentParseSession } from './parsingSession';
export type { OfficeDocumentParseSession } from './parsingSession';
export type {
  OfficeCapabilities,
  OfficeChangeOperation,
  OfficeChangeSet,
  OfficeDocumentMetadata,
  OfficeDocumentModel,
  OfficeDocumentRuntime,
  OfficeDocumentSnapshot,
  OfficeDocumentSource,
  OfficeFormatId,
  OfficeJsonValue,
  OfficeNodeId,
  OfficeResourceLocator,
  OfficeResourceRef,
  OfficeSourceLocation,
  OfficeSourceMap,
  OfficeWarning,
} from './types';

export {
  OfficeFileViewerError,
  OfficeResourceLimitError,
  createOfficeParseSession,
  isOfficeFileViewerError,
} from '../services/parsing';
export type {
  OfficeFileViewerErrorCode,
  OfficeFileViewerErrorContext,
  OfficeFileViewerErrorStage,
  OfficeParseOptions,
  OfficeParseResourcePolicy,
  OfficeParseSession,
  OfficeParseSessionStatus,
  OfficePreviewReadyInfo,
  ParseProgress,
  ParseStage,
  WorkerMode,
} from '../services/parsing';

export { createOfficeResourceStore } from '../services/resource-store/OfficeResourceStore';
export type {
  OfficeResourceSource,
  OfficeResourceStore,
  OfficeResourceStoreOptions,
} from '../services/resource-store/types';
export {
  OfficeExportService,
  createOfficeExportService,
  createOfficeExporterRegistry,
  exportOriginalOfficeFile,
} from './export';
export type {
  OfficeExportError,
  OfficeExportErrorCode,
  OfficeExportRequest,
  OfficeExportResult,
  OfficeExporter,
  OfficeExporterRegistry,
  OfficeOriginalSource,
} from './export';
export {
  createOfficeDocumentSession,
  disposeDocumentSession,
} from './lifecycle';
export type { OfficeDocumentSession, OfficeSessionResource } from './lifecycle';
export { createOfficeResourceSession } from './resources';
export type {
  OfficeResourceResolver,
  OfficeResourceSession,
} from './resources';
