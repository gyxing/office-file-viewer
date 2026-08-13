import { openCfbRandomAccess } from '../../shared/binary/cfb';
import { createBlobRandomAccessSource } from '../../shared/io';
import type { DocBinarySource } from './DocBinarySource';
import { isDocObjectPreviewStreamPath } from './docObjectPreview';
import { readDocTableStreamName } from './readDocStructure';

/** 识别 DOC FIB 结构时预读的字节数。 */
const DOC_FIB_PROBE_BYTES = 2048;

/** 从 Blob 创建只读取 DOC 核心流的 CFB 随机数据源。 */
export async function createCfbDocBinarySource(
  file: Blob,
  signal?: AbortSignal,
): Promise<DocBinarySource> {
  const randomSource = createBlobRandomAccessSource(file);
  const reader = await openCfbRandomAccess(randomSource, signal);

  try {
    const wordDocument = reader.openStream('WordDocument');
    if (!wordDocument) {
      throw new Error('DOC 文件缺少 WordDocument 数据流');
    }
    const fibBase = await wordDocument.read(
      0,
      Math.min(DOC_FIB_PROBE_BYTES, wordDocument.entry.streamSize),
      signal,
    );
    const tableStreamName = readDocTableStreamName(fibBase);
    const table = reader.openStream(tableStreamName);
    if (!table) {
      throw new Error(`DOC 文件缺少 ${tableStreamName} 数据流`);
    }
    const data = reader.openStream('Data');
    const objectPreviewStreams = reader.entries
      .filter(
        (entry) =>
          entry.objectType === 'stream' &&
          isDocObjectPreviewStreamPath(entry.path),
      )
      .flatMap((entry) => {
        const stream = reader.openStream(entry.path);
        return stream ? [{ path: entry.path, stream }] : [];
      });
    let disposed = false;

    return {
      fileSize: file.size,
      wordDocumentSize: wordDocument.entry.streamSize,
      tableStreamName,
      tableSize: table.entry.streamSize,
      dataSize: data?.entry.streamSize ?? 0,
      readWordDocument: (offset, length, readSignal) =>
        wordDocument.read(offset, length, readSignal),
      readTable: (offset, length, readSignal) =>
        table.read(offset, length, readSignal),
      readData: (offset, length, readSignal) => {
        if (!data) {
          if (offset === 0 && length === 0) {
            return Promise.resolve(new Uint8Array());
          }
          throw new Error('DOC 文件缺少 Data 数据流');
        }
        return data.read(offset, length, readSignal);
      },
      async readObjectPreviewStreams(readSignal) {
        const previews: Array<readonly [string, Uint8Array]> = [];
        // 顺序读取可限制复杂文档一次性复制多个大型 EMF 时的瞬时内存。
        for (const { path, stream } of objectPreviewStreams) {
          previews.push([path, await stream.materialize(readSignal)]);
        }
        return previews;
      },
      async dispose() {
        if (disposed) return;
        disposed = true;
        await reader.close();
      },
    };
  } catch (error) {
    await reader.close();
    throw error;
  }
}
