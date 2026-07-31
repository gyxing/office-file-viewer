/** XLS 解析失败原因的稳定代码。 */
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
      /** 发生错误的 XLS 二进制记录标识。 */
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
