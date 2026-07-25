/** 描述 PptParseErrorCode 在 PPT 二进制解析中的数据结构。 */
export type PptParseErrorCode =
  | 'PPT_RECORD_TRUNCATED'
  | 'PPT_RECORD_OUT_OF_RANGE'
  | 'PPT_INVALID_RECORD'
  | 'PPT_REQUIRED_STREAM_MISSING'
  | 'PPT_ENCRYPTED'
  | 'PPT_EDIT_CHAIN_CYCLE'
  | 'PPT_EDIT_CHAIN_INVALID'
  | 'PPT_PERSIST_DIRECTORY_INVALID'
  | 'PPT_DOCUMENT_MISSING'
  | 'PPT_NO_VALID_SLIDES';

/** PPT 解析错误只保留安全的记录上下文，不暴露本地文件路径。 */
export class PptParseError extends Error {
  readonly code: PptParseErrorCode;
  readonly offset?: number;
  readonly recordType?: number;

  constructor(
    code: PptParseErrorCode,
    message: string,
    options: {
      /** PptParseError 在源二进制流中的字节偏移。 */
      offset?: number;
      /** 当前局部结构 从源格式读取的 recordType 枚举或标识值。 */
      recordType?: number;
    } = {},
  ) {
    super(message);
    this.name = 'PptParseError';
    this.code = code;
    this.offset = options.offset;
    this.recordType = options.recordType;
  }
}
