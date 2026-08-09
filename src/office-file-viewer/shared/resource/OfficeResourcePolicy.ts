import { OfficeResourceLimitError } from './OfficeResourceLimitError';

/** 控制 OOXML 归档读取资源上限；未配置的字段保持无限制。 */
export type OfficeArchiveResourcePolicy = {
  /** 允许读取的最大归档条目数量。 */
  maxArchiveEntries?: number;
  /** 单个归档条目允许解压的最大字节数。 */
  maxArchiveEntryBytes?: number;
  /** 整个归档允许解压的累计最大字节数。 */
  maxArchiveInflatedBytes?: number;
  /** 单条目或整个归档允许的最大压缩比。 */
  maxCompressionRatio?: number;
};

/** 控制单次 Office 文件解析使用的资源上限；默认不限制。 */
export type OfficeParseResourcePolicy = OfficeArchiveResourcePolicy & {
  /** 允许解析的最大源文件字节数。 */
  maxFileBytes?: number;
  /** 允许单次解析持续的最大毫秒数。 */
  timeoutMs?: number;
};

/** 资源策略中需要统一校验的数值字段。 */
const OFFICE_RESOURCE_POLICY_FIELDS: ReadonlyArray<
  keyof OfficeParseResourcePolicy
> = [
  'maxFileBytes',
  'maxArchiveEntries',
  'maxArchiveEntryBytes',
  'maxArchiveInflatedBytes',
  'maxCompressionRatio',
  'timeoutMs',
];

/** 校验宿主配置的限制均为大于零的有限数，避免错误配置静默生效。 */
export function validateOfficeResourcePolicy(
  policy?: OfficeParseResourcePolicy,
): void {
  if (!policy) return;
  for (const field of OFFICE_RESOURCE_POLICY_FIELDS) {
    const value = policy[field];
    if (value === undefined) continue;
    if (!Number.isFinite(value) || value <= 0) {
      throw new OfficeResourceLimitError(
        'INVALID_RESOURCE_POLICY',
        `资源策略 ${field} 必须是大于零的有限数`,
        { actual: value },
      );
    }
  }
}
