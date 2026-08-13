import { createCfbDocBinarySource } from './createCfbDocBinarySource';
import {
  parseDocCore,
  type DocCoreContext,
  type DocCoreResult,
} from './parseDocCore';

/** 使用 CFB 随机 Reader 读取大 DOC/WPS，避免先复制完整文件。 */
export async function parseDocRandomAccess(
  file: Blob,
  context: DocCoreContext,
  signal?: AbortSignal,
): Promise<DocCoreResult> {
  const source = await createCfbDocBinarySource(file, signal);
  try {
    const [wordDocument, tableStream, dataStream, objectPreviewStreams] =
      await Promise.all([
        source.readWordDocument(0, source.wordDocumentSize, signal),
        source.readTable(0, source.tableSize, signal),
        source.readData(0, source.dataSize, signal),
        source.readObjectPreviewStreams(signal),
      ]);
    return await parseDocCore(
      {
        wordDocument,
        tableStream,
        imageStreams: [
          ['WordDocument', wordDocument],
          [source.tableStreamName, tableStream],
          ['Data', dataStream],
          ...objectPreviewStreams,
        ],
      },
      context,
    );
  } finally {
    await source.dispose();
  }
}
