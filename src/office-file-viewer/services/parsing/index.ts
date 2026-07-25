// 文件解析层公共入口，导出解析会话工厂及其控制和进度类型。
export { createOfficeParseSession } from './createParseSession';
export type {
  OfficeParseOptions,
  OfficeParseSession,
  OfficeParseSessionStatus,
  ParseProgress,
  ParseStage,
  WorkerMode,
} from './types';
