/** 解析资源策略可以识别的稳定错误代码。 */
export type OfficeResourceLimitCode =
  | 'INVALID_RESOURCE_POLICY'
  | 'FILE_SIZE_LIMIT_EXCEEDED'
  | 'ARCHIVE_ENTRY_COUNT_LIMIT_EXCEEDED'
  | 'ARCHIVE_ENTRY_SIZE_LIMIT_EXCEEDED'
  | 'ARCHIVE_TOTAL_SIZE_LIMIT_EXCEEDED'
  | 'ARCHIVE_COMPRESSION_RATIO_LIMIT_EXCEEDED'
  | 'PARSE_TIMEOUT';

/** 资源限制错误携带的可诊断数值。 */
export type OfficeResourceLimitDetails = {
  /** 触发限制的配置值。 */
  limit?: number;
  /** 实际检测到的值。 */
  actual?: number;
  /** 触发限制的归档条目路径。 */
  path?: string;
};

/** 表示文件解析因调用方配置的资源策略而停止。 */
export class OfficeResourceLimitError extends Error {
  /** 供宿主稳定区分资源限制原因的错误代码。 */
  readonly code: OfficeResourceLimitCode;

  /** 当前策略配置的上限值。 */
  readonly limit?: number;

  /** 触发限制时检测到的实际值。 */
  readonly actual?: number;

  /** 触发限制的归档条目路径。 */
  readonly path?: string;

  constructor(
    code: OfficeResourceLimitCode,
    message: string,
    details: OfficeResourceLimitDetails = {},
  ) {
    super(message);
    this.name = 'OfficeResourceLimitError';
    this.code = code;
    this.limit = details.limit;
    this.actual = details.actual;
    this.path = details.path;
  }
}
