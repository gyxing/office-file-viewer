import type { CfbStreamReader } from '../../shared/binary/cfb';
import type { PptEditChain } from './types';

/** PPT 二进制记录头占用的字节数。 */
const PPT_RECORD_HEADER_BYTES = 8;

function createAbortError() {
  const error = new Error('PPT 持久化对象读取已取消');
  error.name = 'AbortError';
  return error;
}

/** 按主流绝对偏移只读取一个完整 PowerPoint 记录。 */
export async function readPptRecordAtOffset(
  stream: CfbStreamReader,
  sourceOffset: number,
  signal?: AbortSignal,
) {
  if (signal?.aborted) throw createAbortError();
  if (sourceOffset + PPT_RECORD_HEADER_BYTES > stream.entry.streamSize) {
    throw new Error(`PPT 记录 ${sourceOffset} 的记录头超出主流边界`);
  }
  const header = await stream.read(
    sourceOffset,
    PPT_RECORD_HEADER_BYTES,
    signal,
  );
  const view = new DataView(
    header.buffer,
    header.byteOffset,
    header.byteLength,
  );
  const payloadLength = view.getUint32(4, true);
  const recordLength = PPT_RECORD_HEADER_BYTES + payloadLength;
  const endOffset = sourceOffset + recordLength;
  if (
    !Number.isSafeInteger(endOffset) ||
    recordLength < PPT_RECORD_HEADER_BYTES ||
    endOffset > stream.entry.streamSize
  ) {
    throw new Error(`PPT 记录 ${sourceOffset} 的记录长度越界`);
  }
  return {
    sourceOffset,
    type: view.getUint16(2, true),
    bytes: await stream.read(sourceOffset, recordLength, signal),
  };
}

/** 按 Persist Directory 的偏移只读取一个完整 PowerPoint 记录。 */
export async function readPptPersistObject(
  stream: CfbStreamReader,
  editChain: PptEditChain,
  persistId: number,
  signal?: AbortSignal,
) {
  const sourceOffset = editChain.persistOffsets.get(persistId);
  return sourceOffset === undefined
    ? undefined
    : readPptRecordAtOffset(stream, sourceOffset, signal);
}

/** 把单个记录的局部缓冲映射为现有同步解析器可消费的 Persist Directory。 */
export function createLocalPptEditChain(
  editChain: PptEditChain,
  persistId: number,
): PptEditChain {
  return {
    ...editChain,
    persistOffsets: new Map([[persistId, 0]]),
    editOffsets: [],
  };
}
