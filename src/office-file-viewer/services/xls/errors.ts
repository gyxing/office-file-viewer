/** 描述 XlsParseErrorCode 在 XLS/BIFF8 解析中的数据结构。 */
export type XlsParseErrorCode =
  | 'INVALID_CFB'
  | 'UNSUPPORTED_BIFF_VERSION'
  | 'ENCRYPTED_FILE'
  | 'MISSING_WORKBOOK_STREAM'
  | 'TRUNCATED_RECORD'
  | 'CORRUPTED_SECTOR_CHAIN'
  | 'INVALID_RECORD_DATA';

/** XLS 解析错误，保留 BIFF 记录上下文供界面展示和排查。 */
export class XlsParseError extends Error {
  readonly code: XlsParseErrorCode;
  readonly offset?: number;
  readonly recordId?: number;

  constructor(
    code: XlsParseErrorCode,
    message: string,
    context: {
      /** XlsParseError 在源二进制流中的字节偏移。 */
      offset?: number;
      /** 当前内联结构 在源文件记录中的数字标识。 */
      recordId?: number;
    } = {},
  ) {
    super(message);
    this.name = 'XlsParseError';
    this.code = code;
    this.offset = context.offset;
    this.recordId = context.recordId;
  }
}
