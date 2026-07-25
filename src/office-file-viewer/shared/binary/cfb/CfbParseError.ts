/** 描述 CfbParseErrorCode 在 CFB 复合文档中的数据结构。 */
export type CfbParseErrorCode =
  | 'INVALID_SIGNATURE'
  | 'INVALID_HEADER'
  | 'SECTOR_OUT_OF_RANGE'
  | 'CHAIN_CYCLE'
  | 'CHAIN_TRUNCATED'
  | 'DIRECTORY_CORRUPTED';

/** 表示可安全展示的 CFB 结构错误，不包含本地文件路径。 */
export class CfbParseError extends Error {
  readonly code: CfbParseErrorCode;
  readonly sector?: number;
  readonly directoryId?: number;

  constructor(
    code: CfbParseErrorCode,
    message: string,
    context: {
      /** 当前局部结构 在 CFB 扇区链中的扇区索引；未提供时使用来源格式或渲染器的默认行为。 */
      sector?: number;
      /** 当前内联结构 在源文件记录中的数字标识。 */
      directoryId?: number;
    } = {},
  ) {
    super(message);
    this.name = 'CfbParseError';
    this.code = code;
    this.sector = context.sector;
    this.directoryId = context.directoryId;
  }
}
