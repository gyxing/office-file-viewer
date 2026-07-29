import type { OfficeContentStoreError } from './types';

/** 创建可由格式 Source 稳定识别的 ContentStore 错误。 */
export function createContentStoreError(
  code: OfficeContentStoreError['code'],
  message: string,
): OfficeContentStoreError {
  const error = new Error(message) as OfficeContentStoreError;
  error.name = 'OfficeContentStoreError';
  error.code = code;
  return error;
}

/** 创建跨浏览器一致的内容读取取消错误。 */
export function createContentAbortError() {
  const error = new Error('内容读取已取消');
  error.name = 'AbortError';
  return error;
}
