import type { CfbStreamReader } from '../../../shared/binary/cfb';
import { XlsParseError } from '../errors';
import type { Biff8SheetDescriptor, Biff8WorkbookGlobals } from '../types';
import { createParseYieldState, type Biff8Record } from './Biff8Reader';
import {
  parseBiff8WorksheetFromCursor,
  type Biff8AsyncRecordCursor,
} from './parseWorksheet';

/** 通过 CFB StreamReader 每次只读取一个 BIFF 记录。 */
function createStreamRecordCursor(
  stream: CfbStreamReader,
  startOffset: number,
  endOffset: number,
  signal?: AbortSignal,
): Biff8AsyncRecordCursor {
  let cursor = startOffset;
  let pending: Biff8Record | undefined;

  const readNext = async () => {
    if (cursor === endOffset) return undefined;
    if (cursor + 4 > endOffset) {
      throw new XlsParseError('TRUNCATED_RECORD', 'BIFF 记录头被截断', {
        offset: cursor - startOffset,
      });
    }
    const header = await stream.read(cursor, 4, signal);
    const id = header[0] | (header[1] << 8);
    const size = header[2] | (header[3] << 8);
    const dataOffset = cursor + 4;
    if (dataOffset + size > endOffset) {
      throw new XlsParseError('TRUNCATED_RECORD', 'BIFF 记录负载被截断', {
        offset: cursor - startOffset,
        recordId: id,
      });
    }
    const data = await stream.read(dataOffset, size, signal);
    const record = {
      id,
      offset: cursor - startOffset,
      dataOffset: dataOffset - startOffset,
      size,
      data,
    } satisfies Biff8Record;
    cursor = dataOffset + size;
    return record;
  };

  return {
    async peek() {
      pending ??= await readNext();
      return pending;
    },
    async next() {
      if (pending) {
        const record = pending;
        pending = undefined;
        return record;
      }
      return readNext();
    },
  };
}

/** 从单个 Sheet 的 CFB 字节范围解析 BIFF8 Worksheet。 */
export async function parseBiff8WorksheetChunks(
  stream: CfbStreamReader,
  descriptor: Biff8SheetDescriptor,
  endOffset: number,
  globals: Biff8WorkbookGlobals,
  signal?: AbortSignal,
) {
  const startOffset = descriptor.streamOffset;
  const localDescriptor = { ...descriptor, streamOffset: 0 };
  return parseBiff8WorksheetFromCursor(
    createStreamRecordCursor(stream, startOffset, endOffset, signal),
    localDescriptor,
    globals,
    Math.max(0, endOffset - startOffset),
    createParseYieldState(8, async () => {
      if (!signal?.aborted) return;
      const error = new Error('XLS Sheet 读取已取消');
      error.name = 'AbortError';
      throw error;
    }),
  );
}
