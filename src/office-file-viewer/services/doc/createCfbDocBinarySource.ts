import { openCfbRandomAccess } from '../../shared/binary/cfb';
import { createBlobRandomAccessSource } from '../../shared/io';
import type { DocBinarySource } from './DocBinarySource';
import { readDocTableStreamName } from './readDocStructure';

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
