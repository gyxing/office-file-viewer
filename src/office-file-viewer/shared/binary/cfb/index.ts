// CFB 复合文档公共入口，供 DOC、XLS 和 PPT 二进制解析器共享。
export { CfbParseError, type CfbParseErrorCode } from './CfbParseError';
export { CFB_SIGNATURE } from './constants';
export { parseCfb } from './parseCfb';
export type {
  CfbDirectoryEntry,
  CfbFile,
  CfbObjectType,
  CfbReadOptions,
} from './types';
