import {
  openCfbRandomAccess,
  type CfbDirectoryEntry,
  type CfbRandomAccessReader,
  type CfbStreamReader,
} from '../../shared/binary/cfb';
import { createBlobRandomAccessSource } from '../../shared/io/createBlobRandomAccessSource';
import {
  createSpreadsheetPerformanceProfile,
  type SpreadsheetPerformanceProfile,
} from '../spreadsheet/spreadsheetPerformance';
import { createParseYieldState } from './biff8/Biff8Reader';
import { createBiff8SharedStringSource } from './biff8/Biff8SharedStringSource';
import { BIFF8_RECORD } from './biff8/constants';
import { parseBiff8Globals } from './biff8/parseGlobals';
import type { Biff8SheetDescriptor, Biff8WorkbookGlobals } from './types';

/** XLS Source 使用的 Sheet 子流描述符。 */
export type XlsSheetDescriptor = Biff8SheetDescriptor & {
  endOffset: number;
  rowCount: number;
  columnCount: number;
  revision: number;
  status: 'estimated' | 'ready' | 'error';
  errorMessage?: string;
  performance: SpreadsheetPerformanceProfile;
};

/** XLS 大文件共享的随机访问结构。 */
export type XlsStructure = {
  sessionId: string;
  reader: CfbRandomAccessReader;
  workbookStream: CfbStreamReader;
  globals: Biff8WorkbookGlobals;
  descriptors: XlsSheetDescriptor[];
  sharedStrings: ReturnType<typeof createBiff8SharedStringSource>;
  fileSize: number;
  mainStreamSize: number;
};

export type ProfiledXlsArchive = {
  reader: CfbRandomAccessReader;
  fileSize: number;
  mainStreamSize: number;
};

function findStream(entries: readonly CfbDirectoryEntry[], name: string) {
  return entries.find(
    (entry) =>
      entry.objectType === 'stream' &&
      entry.name.toLowerCase() === name.toLowerCase(),
  );
}

/** 使用 CFB 目录提前判断 XLS 是否可能进入大文件路径。 */
export async function profileXlsArchive(
  file: File,
  signal?: AbortSignal,
): Promise<ProfiledXlsArchive> {
  const reader = await openCfbRandomAccess(
    createBlobRandomAccessSource(file),
    signal,
  );
  try {
    return {
      reader,
      fileSize: file.size,
      mainStreamSize:
        findStream(reader.entries, 'Workbook')?.streamSize ??
        findStream(reader.entries, 'Book')?.streamSize ??
        0,
    };
  } catch (error) {
    await reader.close();
    throw error;
  }
}

function concatChunks(chunks: readonly Uint8Array[]) {
  const result = new Uint8Array(
    chunks.reduce((total, chunk) => total + chunk.length, 0),
  );
  let offset = 0;
  chunks.forEach((chunk) => {
    result.set(chunk, offset);
    offset += chunk.length;
  });
  return result;
}

/** 只读取 Workbook Globals 子流，不触碰任何 Sheet 正文。 */
async function readGlobalsBytes(stream: CfbStreamReader, signal?: AbortSignal) {
  const chunks: Uint8Array[] = [];
  let offset = 0;
  while (offset + 4 <= stream.entry.streamSize) {
    const header = await stream.read(offset, 4, signal);
    const view = new DataView(
      header.buffer,
      header.byteOffset,
      header.byteLength,
    );
    const id = view.getUint16(0, true);
    const size = view.getUint16(2, true);
    const data = await stream.read(offset + 4, size, signal);
    chunks.push(header, data);
    offset += 4 + size;
    if (id === BIFF8_RECORD.EOF) break;
    if (chunks.length % 128 === 0) {
      await new Promise<void>((resolve) => {
        setTimeout(resolve, 0);
      });
    }
  }
  return concatChunks(chunks);
}

async function estimateSheetDimensions(
  stream: CfbStreamReader,
  descriptor: Biff8SheetDescriptor,
  endOffset: number,
  signal?: AbortSignal,
) {
  const length = Math.min(128 * 1024, endOffset - descriptor.streamOffset);
  const bytes = await stream.read(descriptor.streamOffset, length, signal);
  let offset = 0;
  while (offset + 4 <= bytes.length) {
    const view = new DataView(
      bytes.buffer,
      bytes.byteOffset + offset,
      bytes.byteLength - offset,
    );
    const id = view.getUint16(0, true);
    const size = view.getUint16(2, true);
    if (offset + 4 + size > bytes.length) break;
    if (id === BIFF8_RECORD.DIMENSIONS && size >= 12) {
      const data = new DataView(
        bytes.buffer,
        bytes.byteOffset + offset + 4,
        size,
      );
      return {
        rowCount: Math.max(1, data.getUint32(4, true)),
        columnCount: Math.max(1, data.getUint16(10, true)),
      };
    }
    offset += 4 + size;
  }
  return { rowCount: 1, columnCount: 1 };
}

/** 读取 Globals、BoundSheet 和维度估算，保持 Sheet 子流未物化。 */
export async function readXlsStructure(
  archive: ProfiledXlsArchive,
  sessionId: string,
  signal?: AbortSignal,
): Promise<{ structure: XlsStructure; requiresSource: boolean }> {
  const workbookStream = archive.reader.openStream('Workbook', 'Book');
  if (!workbookStream) {
    await archive.reader.close();
    throw new Error('XLS 文件缺少 Workbook 数据流');
  }
  const globalsBytes = await readGlobalsBytes(workbookStream, signal);
  const globals = await parseBiff8Globals(
    globalsBytes,
    createParseYieldState(),
    {
      streamLength: workbookStream.entry.streamSize,
      validateSheetBof: false,
    },
  );
  const ordered = [...globals.sheets].sort(
    (left, right) => left.streamOffset - right.streamOffset,
  );
  const dimensions = await Promise.all(
    ordered.map((descriptor, index) =>
      estimateSheetDimensions(
        workbookStream,
        descriptor,
        ordered[index + 1]?.streamOffset ?? workbookStream.entry.streamSize,
        signal,
      ),
    ),
  );
  const descriptors = ordered.map((descriptor, index): XlsSheetDescriptor => {
    const endOffset =
      ordered[index + 1]?.streamOffset ?? workbookStream.entry.streamSize;
    const dimension = dimensions[index];
    return {
      ...descriptor,
      endOffset,
      rowCount: dimension.rowCount,
      columnCount: dimension.columnCount,
      revision: 1,
      status: 'estimated',
      performance: createSpreadsheetPerformanceProfile({
        rowCount: dimension.rowCount,
        columnCount: dimension.columnCount,
        cfbFileBytes: archive.fileSize,
        cfbMainStreamBytes: archive.mainStreamSize,
        sheetBytes: endOffset - descriptor.streamOffset,
      }),
    };
  });
  return {
    structure: {
      sessionId,
      reader: archive.reader,
      workbookStream,
      globals,
      descriptors,
      sharedStrings: createBiff8SharedStringSource(globals.sharedStrings),
      fileSize: archive.fileSize,
      mainStreamSize: archive.mainStreamSize,
    },
    requiresSource: descriptors.some(
      (descriptor) =>
        descriptor.performance.sheetMode === 'lazy' ||
        descriptor.performance.gridMode !== 'table',
    ),
  };
}
