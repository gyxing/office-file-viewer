// PPT 二进制解析公共入口，导出解析、模型适配和结构化错误。
export { adaptPptDocument } from './adapter';
export { PptParseError, type PptParseErrorCode } from './errors';
export { parsePpt } from './parsePpt';
