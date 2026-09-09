/** Core 解析门面，隐藏内部 RuntimeSink、Worker 消息和 Viewer 预览句柄。 */
export type {
  OfficeParseOptions,
  OfficeParseSession,
  OfficeParseSessionStatus,
  OfficePreviewReadyInfo,
  ParseProgress,
  ParseStage,
  WorkerMode,
} from '../services/parsing';
export { createOfficeParseSession } from '../services/parsing/createParseSession';
export { parseOfficeFile } from '../services/preview';
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
