// 文件解析层公共入口，导出解析会话工厂及其控制和进度类型。
export { OfficeResourceLimitError } from '../../shared/resource/OfficeResourceLimitError';
export type { OfficeResourceLimitCode } from '../../shared/resource/OfficeResourceLimitError';
export {
  OfficeFileViewerError,
  isOfficeFileViewerError,
  normalizeOfficeFileViewerError,
} from '../errors/OfficeFileViewerError';
export type {
  OfficeFileViewerErrorCode,
  OfficeFileViewerErrorContext,
  OfficeFileViewerErrorStage,
} from '../errors/OfficeFileViewerError';
export { createOfficeParseSession } from './createParseSession';
export type {
  OfficeParseOptions,
  OfficeParseResourcePolicy,
  OfficeParseSession,
  OfficeParseSessionStatus,
  OfficePreviewReadyInfo,
  ParseProgress,
  ParseStage,
  WorkerMode,
} from './types';
