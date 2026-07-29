import type {
  RandomAccessSource,
  RandomAccessSourceError,
} from './RandomAccessSource';

function createAbortError() {
  const error = new Error('随机读取已取消');
  error.name = 'AbortError';
  return error;
}

function createRangeError(message: string): RandomAccessSourceError {
  const error = new Error(message) as RandomAccessSourceError;
  error.name = 'RandomAccessSourceError';
  error.code = 'INVALID_RANDOM_ACCESS_RANGE';
  return error;
}

/** 使用 Blob.slice() 创建不会预先物化完整文件的随机读取数据源。 */
export function createBlobRandomAccessSource(blob: Blob): RandomAccessSource {
  let closed = false;

  return {
    size: blob.size,
    async read(offset, length, signal) {
      if (closed) {
        const error = new Error(
          '随机读取数据源已关闭',
        ) as RandomAccessSourceError;
        error.name = 'RandomAccessSourceError';
        error.code = 'RANDOM_ACCESS_SOURCE_CLOSED';
        throw error;
      }
      if (
        !Number.isSafeInteger(offset) ||
        !Number.isSafeInteger(length) ||
        offset < 0 ||
        length < 0 ||
        offset + length > blob.size
      ) {
        throw createRangeError('随机读取范围超出数据源边界');
      }
      if (signal?.aborted) throw createAbortError();
      const buffer = await blob.slice(offset, offset + length).arrayBuffer();
      if (signal?.aborted) throw createAbortError();
      return new Uint8Array(buffer);
    },
    async close() {
      closed = true;
    },
  };
}
