import type { RandomAccessSource } from '../../io';
import { CfbParseError } from './CfbParseError';
import {
  buildCfbMiniSectorChain,
  buildCfbSectorChain,
  parseCfbHeader,
  readCfbStructure,
  validateCfbSectorIndex,
} from './CfbStructure';
import {
  CFB_HEADER_SIZE,
  CFB_MINI_SECTOR_SIZE,
  MINI_STREAM_CUTOFF_SIZE,
} from './constants';
import type {
  CfbDirectoryEntry,
  CfbRandomAccessReader,
  CfbStreamReader,
} from './types';

type CfbReaderError = Error & {
  code: 'CFB_READER_CLOSED' | 'INVALID_CFB_STREAM_RANGE';
};

function createAbortError() {
  const error = new Error('CFB 读取已取消');
  error.name = 'AbortError';
  return error;
}

function createReaderError(
  code: CfbReaderError['code'],
  message: string,
): CfbReaderError {
  const error = new Error(message) as CfbReaderError;
  error.name = 'CfbReaderError';
  error.code = code;
  return error;
}

function ensureSignal(signal?: AbortSignal) {
  if (signal?.aborted) throw createAbortError();
}

function createYieldIfNeeded(signal?: AbortSignal) {
  let deadline = Date.now() + 8;
  return async () => {
    ensureSignal(signal);
    if (Date.now() < deadline) return;
    await new Promise<void>((resolve) => {
      setTimeout(resolve, 0);
    });
    ensureSignal(signal);
    deadline = Date.now() + 8;
  };
}

/** 使用 RandomAccessSource 打开只常驻结构索引的 CFB Reader。 */
export async function openCfbRandomAccess(
  source: RandomAccessSource,
  signal?: AbortSignal,
): Promise<CfbRandomAccessReader> {
  ensureSignal(signal);
  const headerBytes = await source.read(
    0,
    Math.min(source.size, CFB_HEADER_SIZE),
    signal,
  );
  const headerView =
    headerBytes.length >= CFB_HEADER_SIZE
      ? new DataView(
          headerBytes.buffer,
          headerBytes.byteOffset,
          headerBytes.byteLength,
        )
      : undefined;
  const sectorShift = headerView?.getUint16(30, true);
  const declaredSectorSize =
    sectorShift === 9 || sectorShift === 12 ? 2 ** sectorShift : 1;
  // 部分 WPS 文件省略最终扇区的零填充，随机路径用虚拟补零保持与兼容物化路径一致。
  const logicalSize =
    declaredSectorSize > 1 && source.size % declaredSectorSize !== 0
      ? source.size + declaredSectorSize - (source.size % declaredSectorSize)
      : source.size;
  const initialHeader = parseCfbHeader(headerBytes, logicalSize);
  let closed = false;
  let closePromise: Promise<void> | undefined;

  const ensureOpen = (readSignal?: AbortSignal) => {
    if (closed) {
      throw createReaderError('CFB_READER_CLOSED', 'CFB Reader 已关闭');
    }
    ensureSignal(readSignal);
  };
  const readSector = async (sector: number, readSignal?: AbortSignal) => {
    ensureOpen(readSignal);
    validateCfbSectorIndex(sector, initialHeader);
    const offset = (sector + 1) * initialHeader.sectorSize;
    const available = Math.max(
      0,
      Math.min(initialHeader.sectorSize, source.size - offset),
    );
    const data = await source.read(offset, available, readSignal);
    if (available === initialHeader.sectorSize) return data;
    const padded = new Uint8Array(initialHeader.sectorSize);
    padded.set(data);
    return padded;
  };
  const structure = await readCfbStructure(
    headerBytes,
    logicalSize,
    readSector,
    {
      signal,
      yieldIfNeeded: createYieldIfNeeded(signal),
    },
  );
  const { entries, fat, header, miniFat } = structure;
  const root = entries.find((entry) => entry.objectType === 'root')!;
  const rootChain = await buildCfbSectorChain(
    root.startSector,
    fat,
    header,
    { signal, yieldIfNeeded: createYieldIfNeeded(signal) },
    Math.ceil(root.streamSize / header.sectorSize),
  );
  const leafEntries = entries.filter((entry) => entry.objectType === 'stream');
  const streamReaders = new Map<number, CfbStreamReader>();

  const readRegularRange = async (
    chain: readonly number[],
    offset: number,
    length: number,
    readSignal?: AbortSignal,
  ) => {
    const result = new Uint8Array(length);
    const yieldIfNeeded = createYieldIfNeeded(readSignal);
    let sourceOffset = offset;
    let targetOffset = 0;
    while (targetOffset < length) {
      ensureOpen(readSignal);
      const chainIndex = Math.floor(sourceOffset / header.sectorSize);
      const withinSector = sourceOffset % header.sectorSize;
      const sector = chain[chainIndex];
      if (sector === undefined) {
        throw new CfbParseError(
          'CHAIN_TRUNCATED',
          'CFB 扇区链不足以覆盖请求范围',
        );
      }
      const take = Math.min(
        length - targetOffset,
        header.sectorSize - withinSector,
      );
      const sectorBytes = await readSector(sector, readSignal);
      result.set(
        sectorBytes.subarray(withinSector, withinSector + take),
        targetOffset,
      );
      sourceOffset += take;
      targetOffset += take;
      await yieldIfNeeded();
    }
    ensureOpen(readSignal);
    return result;
  };

  const createStreamReader = (entry: CfbDirectoryEntry): CfbStreamReader => {
    let chainPromise: Promise<number[]> | undefined;
    const getChain = (readSignal?: AbortSignal) => {
      if (!chainPromise) {
        const pending =
          entry.streamSize < MINI_STREAM_CUTOFF_SIZE
            ? buildCfbMiniSectorChain(entry, miniFat, root.streamSize, {
                signal: readSignal,
                yieldIfNeeded: createYieldIfNeeded(readSignal),
              })
            : buildCfbSectorChain(
                entry.startSector,
                fat,
                header,
                {
                  signal: readSignal,
                  yieldIfNeeded: createYieldIfNeeded(readSignal),
                },
                Math.ceil(entry.streamSize / header.sectorSize),
              );
        chainPromise = pending;
        void pending.catch(() => {
          if (chainPromise === pending) chainPromise = undefined;
        });
      }
      return chainPromise;
    };

    return {
      entry,
      async read(offset, length, readSignal) {
        ensureOpen(readSignal);
        if (
          !Number.isSafeInteger(offset) ||
          !Number.isSafeInteger(length) ||
          offset < 0 ||
          length < 0 ||
          offset + length > entry.streamSize
        ) {
          throw createReaderError(
            'INVALID_CFB_STREAM_RANGE',
            `CFB 流 ${entry.path} 的读取范围无效`,
          );
        }
        if (length === 0) return new Uint8Array();
        const chain = await getChain(readSignal);
        if (entry.streamSize >= MINI_STREAM_CUTOFF_SIZE) {
          return readRegularRange(chain, offset, length, readSignal);
        }

        const result = new Uint8Array(length);
        const yieldIfNeeded = createYieldIfNeeded(readSignal);
        let sourceOffset = offset;
        let targetOffset = 0;
        while (targetOffset < length) {
          ensureOpen(readSignal);
          const chainIndex = Math.floor(sourceOffset / CFB_MINI_SECTOR_SIZE);
          const withinMiniSector = sourceOffset % CFB_MINI_SECTOR_SIZE;
          const miniSector = chain[chainIndex];
          if (miniSector === undefined) {
            throw new CfbParseError(
              'CHAIN_TRUNCATED',
              `CFB 小流 ${entry.path} 的扇区链不足`,
              { directoryId: entry.id },
            );
          }
          const take = Math.min(
            length - targetOffset,
            CFB_MINI_SECTOR_SIZE - withinMiniSector,
          );
          const chunk = await readRegularRange(
            rootChain,
            miniSector * CFB_MINI_SECTOR_SIZE + withinMiniSector,
            take,
            readSignal,
          );
          result.set(chunk, targetOffset);
          sourceOffset += take;
          targetOffset += take;
          await yieldIfNeeded();
        }
        return result;
      },
      materialize(readSignal) {
        return this.read(0, entry.streamSize, readSignal);
      },
    };
  };

  return {
    entries,
    hasEntry(name) {
      const normalized = name.toLowerCase();
      return entries.some(
        (entry) =>
          entry.path.toLowerCase() === normalized ||
          entry.name.toLowerCase() === normalized,
      );
    },
    openStream(...names) {
      ensureOpen();
      let target: CfbDirectoryEntry | undefined;
      for (const name of names) {
        const exactPath = name.startsWith('/') ? name : `/${name}`;
        target = leafEntries.find(
          (entry) => entry.path === name || entry.path === exactPath,
        );
        if (!target) {
          const matches = leafEntries.filter(
            (entry) => entry.name.toLowerCase() === name.toLowerCase(),
          );
          if (matches.length === 1) target = matches[0];
        }
        if (target) break;
      }
      if (!target) return undefined;
      const cached = streamReaders.get(target.id);
      if (cached) return cached;
      const reader = createStreamReader(target);
      streamReaders.set(target.id, reader);
      return reader;
    },
    close() {
      if (closePromise) return closePromise;
      closed = true;
      streamReaders.clear();
      closePromise = source.close();
      return closePromise;
    },
  };
}
