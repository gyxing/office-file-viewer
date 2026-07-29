import { CfbParseError } from './CfbParseError';
import {
  CFB_DIRECTORY_ENTRY_SIZE,
  CFB_HEADER_SIZE,
  CFB_MINI_SECTOR_SIZE,
  CFB_SIGNATURE,
  DIFAT_SECTOR,
  END_OF_CHAIN,
  FAT_SECTOR,
  FREE_SECTOR,
  MINI_STREAM_CUTOFF_SIZE,
  NO_STREAM,
} from './constants';
import type { CfbDirectoryEntry, CfbObjectType, CfbReadOptions } from './types';

/** 描述 CFB Header 中后续随机读取必需的结构信息。 */
export type CfbHeader = {
  majorVersion: number;
  sectorSize: number;
  sectorCount: number;
  fatSectorCount: number;
  directoryStartSector: number;
  directorySectorCount: number;
  miniFatStartSector: number;
  miniFatSectorCount: number;
  difatStartSector: number;
  difatSectorCount: number;
};

type RawDirectoryEntry = Omit<CfbDirectoryEntry, 'path'>;

export type CfbSectorReader = (
  sector: number,
  signal?: AbortSignal,
) => Promise<Uint8Array>;

/** 保存随机 Reader 和完整物化路径共用的 CFB 结构索引。 */
export type CfbStructure = {
  header: CfbHeader;
  fat: number[];
  miniFat: number[];
  entries: CfbDirectoryEntry[];
};

function readUint16(view: DataView, offset: number) {
  return view.getUint16(offset, true);
}

function readUint32(view: DataView, offset: number) {
  return view.getUint32(offset, true);
}

function createAbortError() {
  const error = new Error('CFB 读取已取消');
  error.name = 'AbortError';
  return error;
}

function ensureNotAborted(signal?: AbortSignal) {
  if (signal?.aborted) throw createAbortError();
}

async function checkpoint(options: CfbReadOptions) {
  ensureNotAborted(options.signal);
  await options.yieldIfNeeded?.();
  ensureNotAborted(options.signal);
}

function isSpecialSector(value: number) {
  return (
    value === FREE_SECTOR ||
    value === END_OF_CHAIN ||
    value === FAT_SECTOR ||
    value === DIFAT_SECTOR
  );
}

/** 校验普通扇区索引，供结构读取和流范围读取复用。 */
export function validateCfbSectorIndex(sector: number, header: CfbHeader) {
  if (!Number.isInteger(sector) || sector < 0 || sector >= header.sectorCount) {
    throw new CfbParseError(
      'SECTOR_OUT_OF_RANGE',
      `CFB 扇区 ${sector} 超出有效范围`,
      { sector },
    );
  }
}

/** 兼容省略最后扇区零填充的 WPS 文件，并保留严格结构校验。 */
export function normalizeCfbBytes(
  source: Uint8Array,
  allowPartialFinalSector = false,
) {
  if (!allowPartialFinalSector || source.length < CFB_HEADER_SIZE)
    return source;
  const view = new DataView(
    source.buffer,
    source.byteOffset,
    source.byteLength,
  );
  const sectorShift = readUint16(view, 30);
  if (sectorShift !== 9 && sectorShift !== 12) return source;
  const sectorSize = 2 ** sectorShift;
  const remainder = source.length % sectorSize;
  if (!remainder) return source;

  const padded = new Uint8Array(source.length + sectorSize - remainder);
  padded.set(source);
  return padded;
}

/** 解析并校验 CFB Header，但不读取正文扇区。 */
export function parseCfbHeader(
  headerBytes: Uint8Array,
  fileSize: number,
): CfbHeader {
  if (headerBytes.length < CFB_HEADER_SIZE) {
    throw new CfbParseError('INVALID_HEADER', 'CFB Header 长度不足');
  }
  if (!CFB_SIGNATURE.every((value, index) => headerBytes[index] === value)) {
    throw new CfbParseError('INVALID_SIGNATURE', '不是有效的 CFB 文件');
  }

  const view = new DataView(
    headerBytes.buffer,
    headerBytes.byteOffset,
    headerBytes.byteLength,
  );
  const majorVersion = readUint16(view, 26);
  const byteOrder = readUint16(view, 28);
  const sectorShift = readUint16(view, 30);
  const miniSectorShift = readUint16(view, 32);
  const miniStreamCutoff = readUint32(view, 56);
  const expectedSectorShift = majorVersion === 3 ? 9 : 12;

  if (
    (majorVersion !== 3 && majorVersion !== 4) ||
    byteOrder !== 0xfffe ||
    sectorShift !== expectedSectorShift ||
    miniSectorShift !== 6 ||
    miniStreamCutoff !== MINI_STREAM_CUTOFF_SIZE
  ) {
    throw new CfbParseError('INVALID_HEADER', 'CFB Header 固定字段无效');
  }

  const sectorSize = 2 ** sectorShift;
  if (fileSize < sectorSize || fileSize % sectorSize !== 0) {
    throw new CfbParseError(
      'INVALID_HEADER',
      'CFB 文件长度与声明的扇区大小不一致',
    );
  }

  const header: CfbHeader = {
    majorVersion,
    sectorSize,
    sectorCount: fileSize / sectorSize - 1,
    directorySectorCount: readUint32(view, 40),
    fatSectorCount: readUint32(view, 44),
    directoryStartSector: readUint32(view, 48),
    miniFatStartSector: readUint32(view, 60),
    miniFatSectorCount: readUint32(view, 64),
    difatStartSector: readUint32(view, 68),
    difatSectorCount: readUint32(view, 72),
  };

  if (
    header.sectorCount < 1 ||
    (majorVersion === 3 && header.directorySectorCount !== 0) ||
    header.fatSectorCount > header.sectorCount ||
    header.miniFatSectorCount > header.sectorCount ||
    header.difatSectorCount > header.sectorCount
  ) {
    throw new CfbParseError('INVALID_HEADER', 'CFB Header 扇区计数无效');
  }
  validateCfbSectorIndex(header.directoryStartSector, header);
  return header;
}

/** 解析 FAT 链并返回经过统一成环、越界和长度校验的扇区索引。 */
export async function buildCfbSectorChain(
  startSector: number,
  fat: readonly number[],
  header: CfbHeader,
  options: CfbReadOptions = {},
  expectedSectors?: number,
) {
  if (startSector === END_OF_CHAIN && (expectedSectors ?? 0) === 0) return [];
  if (isSpecialSector(startSector)) {
    throw new CfbParseError('CHAIN_TRUNCATED', 'CFB 扇区链起点无效', {
      sector: startSector,
    });
  }

  const chain: number[] = [];
  const visited = new Set<number>();
  let sector = startSector;
  while (sector !== END_OF_CHAIN) {
    validateCfbSectorIndex(sector, header);
    if (visited.has(sector)) {
      throw new CfbParseError('CHAIN_CYCLE', `CFB 扇区链在 ${sector} 成环`, {
        sector,
      });
    }
    if (sector >= fat.length) {
      throw new CfbParseError(
        'CHAIN_TRUNCATED',
        `CFB FAT 缺少扇区 ${sector} 的链项`,
        { sector },
      );
    }
    visited.add(sector);
    chain.push(sector);
    const next = fat[sector];
    if (next !== END_OF_CHAIN && isSpecialSector(next)) {
      throw new CfbParseError(
        'CHAIN_TRUNCATED',
        `CFB 扇区 ${sector} 指向无效标记`,
        { sector },
      );
    }
    sector = next;
    await checkpoint(options);
  }
  if (expectedSectors !== undefined && chain.length !== expectedSectors) {
    throw new CfbParseError(
      'CHAIN_TRUNCATED',
      `CFB 扇区链长度 ${chain.length} 与声明值 ${expectedSectors} 不一致`,
    );
  }
  return chain;
}

/** 按统一扇区链规则物化一段完整 CFB 流。 */
export async function readCfbSectorChain(
  startSector: number,
  fat: readonly number[],
  header: CfbHeader,
  readSector: CfbSectorReader,
  options: CfbReadOptions,
  expectedSectors?: number,
) {
  const chain = await buildCfbSectorChain(
    startSector,
    fat,
    header,
    options,
    expectedSectors,
  );
  const result = new Uint8Array(chain.length * header.sectorSize);
  for (let index = 0; index < chain.length; index += 1) {
    const data = await readSector(chain[index], options.signal);
    if (data.length !== header.sectorSize) {
      throw new CfbParseError(
        'SECTOR_OUT_OF_RANGE',
        `CFB 扇区 ${chain[index]} 的字节范围无效`,
        { sector: chain[index] },
      );
    }
    result.set(data, index * header.sectorSize);
    await checkpoint(options);
  }
  return result;
}

async function readDifat(
  headerBytes: Uint8Array,
  header: CfbHeader,
  readSector: CfbSectorReader,
  options: CfbReadOptions,
) {
  const headerView = new DataView(
    headerBytes.buffer,
    headerBytes.byteOffset,
    CFB_HEADER_SIZE,
  );
  const difat: number[] = [];
  for (let offset = 76; offset < CFB_HEADER_SIZE; offset += 4) {
    const sector = readUint32(headerView, offset);
    if (sector !== FREE_SECTOR) difat.push(sector);
  }

  const visited = new Set<number>();
  let sector = header.difatStartSector;
  const entriesPerSector = header.sectorSize / 4 - 1;
  for (let index = 0; index < header.difatSectorCount; index += 1) {
    if (sector === END_OF_CHAIN) {
      throw new CfbParseError('CHAIN_TRUNCATED', 'CFB DIFAT 链提前结束');
    }
    validateCfbSectorIndex(sector, header);
    if (visited.has(sector)) {
      throw new CfbParseError('CHAIN_CYCLE', `CFB DIFAT 扇区 ${sector} 成环`, {
        sector,
      });
    }
    visited.add(sector);
    const data = await readSector(sector, options.signal);
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    for (let entry = 0; entry < entriesPerSector; entry += 1) {
      const fatSector = readUint32(view, entry * 4);
      if (fatSector !== FREE_SECTOR) difat.push(fatSector);
    }
    sector = readUint32(view, header.sectorSize - 4);
    await checkpoint(options);
  }

  if (
    (header.difatSectorCount === 0 &&
      header.difatStartSector !== END_OF_CHAIN) ||
    (header.difatSectorCount > 0 && sector !== END_OF_CHAIN)
  ) {
    throw new CfbParseError('INVALID_HEADER', 'CFB DIFAT 链计数不一致');
  }
  if (difat.length < header.fatSectorCount) {
    throw new CfbParseError('CHAIN_TRUNCATED', 'CFB FAT 扇区列表不完整');
  }
  return difat.slice(0, header.fatSectorCount);
}

async function readFat(
  header: CfbHeader,
  difat: readonly number[],
  readSector: CfbSectorReader,
  options: CfbReadOptions,
) {
  const fat: number[] = [];
  const visited = new Set<number>();
  for (const sector of difat) {
    validateCfbSectorIndex(sector, header);
    if (visited.has(sector)) {
      throw new CfbParseError('CHAIN_CYCLE', `CFB FAT 扇区 ${sector} 重复`, {
        sector,
      });
    }
    visited.add(sector);
    const data = await readSector(sector, options.signal);
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    for (let offset = 0; offset < data.length; offset += 4) {
      fat.push(readUint32(view, offset));
    }
    await checkpoint(options);
  }
  return fat;
}

function decodeDirectoryName(bytes: Uint8Array, nameLength: number) {
  if (nameLength < 2 || nameLength > 64 || nameLength % 2 !== 0) {
    throw new CfbParseError('DIRECTORY_CORRUPTED', 'CFB 目录项名称长度无效');
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let result = '';
  for (let offset = 0; offset < nameLength - 2; offset += 2) {
    result += String.fromCharCode(readUint16(view, offset));
  }
  return result;
}

function objectTypeFromValue(value: number): CfbObjectType | undefined {
  if (value === 1) return 'storage';
  if (value === 2) return 'stream';
  if (value === 5) return 'root';
  return undefined;
}

async function parseDirectoryEntries(
  directoryStream: Uint8Array,
  header: CfbHeader,
  options: CfbReadOptions,
) {
  if (directoryStream.length % CFB_DIRECTORY_ENTRY_SIZE !== 0) {
    throw new CfbParseError(
      'DIRECTORY_CORRUPTED',
      'CFB 目录流未按 128 字节对齐',
    );
  }

  const entries: Array<RawDirectoryEntry | undefined> = [];
  for (
    let offset = 0, id = 0;
    offset < directoryStream.length;
    offset += CFB_DIRECTORY_ENTRY_SIZE, id += 1
  ) {
    const data = directoryStream.subarray(
      offset,
      offset + CFB_DIRECTORY_ENTRY_SIZE,
    );
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const rawObjectType = data[66];
    if (rawObjectType === 0) {
      entries.push(undefined);
      continue;
    }
    const objectType = objectTypeFromValue(rawObjectType);
    if (!objectType) {
      throw new CfbParseError(
        'DIRECTORY_CORRUPTED',
        `CFB 目录项 ${id} 的对象类型无效`,
        { directoryId: id },
      );
    }
    const lowSize = readUint32(view, 120);
    const highSize = header.majorVersion === 3 ? 0 : readUint32(view, 124);
    const streamSize = highSize * 0x100000000 + lowSize;
    if (!Number.isSafeInteger(streamSize)) {
      throw new CfbParseError(
        'DIRECTORY_CORRUPTED',
        `CFB 目录项 ${id} 的流长度超出安全范围`,
        { directoryId: id },
      );
    }
    entries.push({
      id,
      name: decodeDirectoryName(data.subarray(0, 64), readUint16(view, 64)),
      objectType,
      startSector: readUint32(view, 116),
      streamSize,
      leftSiblingId: readUint32(view, 68),
      rightSiblingId: readUint32(view, 72),
      childId: readUint32(view, 76),
    });
    await checkpoint(options);
  }
  return entries;
}

function getDirectoryEntry(
  entries: Array<RawDirectoryEntry | undefined>,
  id: number,
) {
  if (id === NO_STREAM) return undefined;
  const entry = entries[id];
  if (!entry) {
    throw new CfbParseError('DIRECTORY_CORRUPTED', `CFB 目录项 ${id} 不存在`, {
      directoryId: id,
    });
  }
  return entry;
}

function assignDirectoryPaths(
  rawEntries: Array<RawDirectoryEntry | undefined>,
) {
  const root = rawEntries.find((entry) => entry?.objectType === 'root');
  if (!root) {
    throw new CfbParseError('DIRECTORY_CORRUPTED', 'CFB 目录缺少 Root Entry');
  }

  const result: CfbDirectoryEntry[] = [{ ...root, path: '/' }];
  const visited = new Set<number>([root.id]);
  const stack: Array<{ id: number; parentPath: string }> = [];
  if (root.childId !== NO_STREAM) {
    stack.push({ id: root.childId, parentPath: '' });
  }

  while (stack.length) {
    const current = stack.pop()!;
    const entry = getDirectoryEntry(rawEntries, current.id)!;
    if (entry.objectType === 'root' || visited.has(entry.id)) {
      throw new CfbParseError(
        'DIRECTORY_CORRUPTED',
        `CFB 目录树在目录项 ${entry.id} 成环`,
        { directoryId: entry.id },
      );
    }
    visited.add(entry.id);
    const path = `${current.parentPath}/${entry.name}`;
    result.push({ ...entry, path });

    if (entry.rightSiblingId !== NO_STREAM) {
      stack.push({
        id: entry.rightSiblingId,
        parentPath: current.parentPath,
      });
    }
    if (entry.leftSiblingId !== NO_STREAM) {
      stack.push({ id: entry.leftSiblingId, parentPath: current.parentPath });
    }
    if (entry.objectType === 'storage' && entry.childId !== NO_STREAM) {
      stack.push({ id: entry.childId, parentPath: path });
    } else if (entry.objectType === 'stream' && entry.childId !== NO_STREAM) {
      throw new CfbParseError(
        'DIRECTORY_CORRUPTED',
        `CFB 流目录项 ${entry.id} 不应包含子项`,
        { directoryId: entry.id },
      );
    }
  }

  const unreachable = rawEntries.filter((entry): entry is RawDirectoryEntry =>
    Boolean(entry && !visited.has(entry.id)),
  );
  if (unreachable.length) {
    throw new CfbParseError(
      'DIRECTORY_CORRUPTED',
      `CFB 目录存在不可达目录项 ${unreachable[0].id}`,
      { directoryId: unreachable[0].id },
    );
  }
  return result.sort((left, right) => left.id - right.id);
}

/** 校验小流链并返回 mini-sector 索引。 */
export async function buildCfbMiniSectorChain(
  entry: CfbDirectoryEntry,
  miniFat: readonly number[],
  miniStreamSize: number,
  options: CfbReadOptions = {},
) {
  if (entry.streamSize === 0) return [];
  const chain: number[] = [];
  const visited = new Set<number>();
  const expectedSectors = Math.ceil(entry.streamSize / CFB_MINI_SECTOR_SIZE);
  let sector = entry.startSector;

  while (sector !== END_OF_CHAIN) {
    if (
      !Number.isInteger(sector) ||
      sector < 0 ||
      sector >= miniFat.length ||
      (sector + 1) * CFB_MINI_SECTOR_SIZE > miniStreamSize
    ) {
      throw new CfbParseError(
        'SECTOR_OUT_OF_RANGE',
        `CFB 小流目录项 ${entry.id} 的扇区 ${sector} 无效`,
        { sector, directoryId: entry.id },
      );
    }
    if (visited.has(sector)) {
      throw new CfbParseError(
        'CHAIN_CYCLE',
        `CFB 小流目录项 ${entry.id} 的扇区链成环`,
        { sector, directoryId: entry.id },
      );
    }
    visited.add(sector);
    chain.push(sector);
    const next = miniFat[sector];
    if (next !== END_OF_CHAIN && isSpecialSector(next)) {
      throw new CfbParseError(
        'CHAIN_TRUNCATED',
        `CFB 小流目录项 ${entry.id} 指向无效标记`,
        { sector, directoryId: entry.id },
      );
    }
    sector = next;
    await checkpoint(options);
  }
  if (chain.length !== expectedSectors) {
    throw new CfbParseError(
      'CHAIN_TRUNCATED',
      `CFB 小流目录项 ${entry.id} 的链长度与流大小不一致`,
      { directoryId: entry.id },
    );
  }
  return chain;
}

/** 只读取 Header、FAT、MiniFAT 和 Directory，建立共享结构索引。 */
export async function readCfbStructure(
  headerBytes: Uint8Array,
  fileSize: number,
  readSector: CfbSectorReader,
  options: CfbReadOptions = {},
): Promise<CfbStructure> {
  ensureNotAborted(options.signal);
  const header = parseCfbHeader(headerBytes, fileSize);
  const difat = await readDifat(headerBytes, header, readSector, options);
  const fat = await readFat(header, difat, readSector, options);
  const directoryStream = await readCfbSectorChain(
    header.directoryStartSector,
    fat,
    header,
    readSector,
    options,
    header.majorVersion === 4 ? header.directorySectorCount : undefined,
  );
  const rawEntries = await parseDirectoryEntries(
    directoryStream,
    header,
    options,
  );
  const entries = assignDirectoryPaths(rawEntries);
  let miniFat: number[] = [];
  if (header.miniFatSectorCount > 0) {
    const data = await readCfbSectorChain(
      header.miniFatStartSector,
      fat,
      header,
      readSector,
      options,
      header.miniFatSectorCount,
    );
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    miniFat = [];
    for (let offset = 0; offset < data.length; offset += 4) {
      miniFat.push(readUint32(view, offset));
    }
  }
  return { header, fat, miniFat, entries };
}
