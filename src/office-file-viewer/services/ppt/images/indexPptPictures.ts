import type { CfbStreamReader } from '../../../shared/binary/cfb';
import { parseOfficeArtRecords } from '../../../shared/officeart';
import type { PptParseContext } from '../types';
import { registerPptPictureRecords } from './readPictures';

/** OfficeArt 记录头占用的字节数。 */
const OFFICE_ART_RECORD_HEADER_BYTES = 8;

function createAbortError() {
  const error = new Error('PPT 图片读取已取消');
  error.name = 'AbortError';
  return error;
}

/** 顺序索引 Pictures 流中的 OfficeArt 记录，避免先物化整个资源流。 */
export async function indexPptPictures(
  stream: CfbStreamReader,
  context: PptParseContext,
  signal?: AbortSignal,
) {
  let offset = 0;
  while (offset < stream.entry.streamSize) {
    if (signal?.aborted) throw createAbortError();
    const remaining = stream.entry.streamSize - offset;
    if (remaining < OFFICE_ART_RECORD_HEADER_BYTES) {
      throw new Error('PPT Pictures 流尾部缺少完整 OfficeArt 记录头');
    }
    const header = await stream.read(
      offset,
      OFFICE_ART_RECORD_HEADER_BYTES,
      signal,
    );
    const view = new DataView(
      header.buffer,
      header.byteOffset,
      header.byteLength,
    );
    const payloadLength = view.getUint32(4, true);
    const recordLength = OFFICE_ART_RECORD_HEADER_BYTES + payloadLength;
    const endOffset = offset + recordLength;
    if (
      !Number.isSafeInteger(endOffset) ||
      recordLength < OFFICE_ART_RECORD_HEADER_BYTES ||
      endOffset > stream.entry.streamSize
    ) {
      throw new Error('PPT Pictures 流中的 OfficeArt 记录长度越界');
    }
    const bytes = await stream.read(offset, recordLength, signal);
    await registerPptPictureRecords(
      parseOfficeArtRecords(bytes, context.warnings),
      context,
    );
    offset = endOffset;
  }
  return context.blipUrls;
}
