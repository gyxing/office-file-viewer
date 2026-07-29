import { CfbParseError } from './CfbParseError';
import {
  buildCfbMiniSectorChain,
  normalizeCfbBytes,
  parseCfbHeader,
  readCfbSectorChain,
  readCfbStructure,
} from './CfbStructure';
import {
  CFB_HEADER_SIZE,
  CFB_MINI_SECTOR_SIZE,
  MINI_STREAM_CUTOFF_SIZE,
} from './constants';
import type { CfbDirectoryEntry, CfbFile, CfbReadOptions } from './types';

function createAbortError() {
  const error = new Error('CFB 读取已取消');
  error.name = 'AbortError';
  return error;
}

function ensureNotAborted(signal?: AbortSignal) {
  if (signal?.aborted) throw createAbortError();
}

/** 创建兼容原有按路径和唯一流名称查询行为的完整 CFB 结果。 */
function createCfbFile(
  entries: CfbDirectoryEntry[],
  streams: Map<string, Uint8Array>,
): CfbFile {
  const leafEntries = entries.filter((entry) => entry.objectType === 'stream');
  return {
    entries,
    streams,
    getStream: (...names) => {
      for (const name of names) {
        const exactPath = name.startsWith('/') ? name : `/${name}`;
        const exact = streams.get(name) ?? streams.get(exactPath);
        if (exact) return exact;
        const matches = leafEntries.filter(
          (entry) => entry.name.toLowerCase() === name.toLowerCase(),
        );
        if (matches.length === 1) return streams.get(matches[0].path);
      }
      return undefined;
    },
    hasEntry: (name) => {
      const normalized = name.toLowerCase();
      return entries.some(
        (entry) =>
          entry.path.toLowerCase() === normalized ||
          entry.name.toLowerCase() === normalized,
      );
    },
  };
}

/** 解析 CFB 容器，并返回可按完整路径或唯一流名称读取的数据流。 */
export async function parseCfb(
  input: ArrayBuffer | Uint8Array,
  options: CfbReadOptions = {},
): Promise<CfbFile> {
  const source = input instanceof Uint8Array ? input : new Uint8Array(input);
  const bytes = normalizeCfbBytes(source, options.allowPartialFinalSector);
  const headerBytes = bytes.subarray(0, CFB_HEADER_SIZE);
  const initialHeader = parseCfbHeader(headerBytes, bytes.length);
  const readSector = async (sector: number) => {
    const offset = (sector + 1) * initialHeader.sectorSize;
    const end = offset + initialHeader.sectorSize;
    if (offset < 0 || end > bytes.length) {
      throw new CfbParseError(
        'SECTOR_OUT_OF_RANGE',
        `CFB 扇区 ${sector} 的字节范围无效`,
        { sector },
      );
    }
    return bytes.subarray(offset, end);
  };
  const structure = await readCfbStructure(
    headerBytes,
    bytes.length,
    readSector,
    options,
  );
  const { entries, fat, header, miniFat } = structure;
  const root = entries.find((entry) => entry.objectType === 'root')!;
  const rootSectorCount = Math.ceil(root.streamSize / header.sectorSize);
  const miniStream =
    root.streamSize > 0
      ? (
          await readCfbSectorChain(
            root.startSector,
            fat,
            header,
            readSector,
            options,
            rootSectorCount,
          )
        ).subarray(0, root.streamSize)
      : new Uint8Array();
  const streams = new Map<string, Uint8Array>();

  for (const entry of entries) {
    if (entry.objectType !== 'stream') continue;
    if (streams.has(entry.path)) {
      throw new CfbParseError(
        'DIRECTORY_CORRUPTED',
        `CFB 目录存在重复路径 ${entry.path}`,
        { directoryId: entry.id },
      );
    }

    let data: Uint8Array;
    if (entry.streamSize === 0) {
      data = new Uint8Array();
    } else if (entry.streamSize < MINI_STREAM_CUTOFF_SIZE) {
      const chain = await buildCfbMiniSectorChain(
        entry,
        miniFat,
        miniStream.length,
        options,
      );
      const materialized = new Uint8Array(chain.length * CFB_MINI_SECTOR_SIZE);
      chain.forEach((sector, index) => {
        const offset = sector * CFB_MINI_SECTOR_SIZE;
        materialized.set(
          miniStream.subarray(offset, offset + CFB_MINI_SECTOR_SIZE),
          index * CFB_MINI_SECTOR_SIZE,
        );
      });
      data = materialized.subarray(0, entry.streamSize);
    } else {
      const sectorCount = Math.ceil(entry.streamSize / header.sectorSize);
      data = (
        await readCfbSectorChain(
          entry.startSector,
          fat,
          header,
          readSector,
          options,
          sectorCount,
        )
      ).subarray(0, entry.streamSize);
    }
    streams.set(entry.path, data);
    ensureNotAborted(options.signal);
    await options.yieldIfNeeded?.();
    ensureNotAborted(options.signal);
  }
  return createCfbFile(entries, streams);
}
