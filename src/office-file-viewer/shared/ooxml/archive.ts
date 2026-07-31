import type {
  OfficeArchiveOpenOptions,
  OfficeArchiveReader,
} from './OfficeArchiveReader';
import { createZipJsOfficeArchiveReader } from './ZipJsOfficeArchiveReader';

export type {
  OfficeArchiveEntry,
  OfficeArchiveOpenOptions,
  OfficeArchiveReader,
} from './OfficeArchiveReader';

/** 打开 OOXML 归档时支持的二进制输入。 */
export type OfficeZipInput = File | Blob | ArrayBuffer | Uint8Array;
/** 按包内路径索引的 OOXML 文本或二进制条目。 */
export type OfficeEntryMap = Map<string, string | Uint8Array>;

/** 检测浏览器是否真正支持 zip.js native 核心需要的 deflate-raw。 */
function supportsNativeZipCore() {
  if (
    typeof document === 'undefined' ||
    typeof DecompressionStream === 'undefined'
  ) {
    return false;
  }
  try {
    new DecompressionStream('deflate-raw');
    return true;
  } catch {
    return false;
  }
}

/**
 * 动态加载匹配当前浏览器能力的 zip.js 核心，并打开 OOXML 归档。
 */
export async function openOfficeArchive(
  file: OfficeZipInput,
  options: OfficeArchiveOpenOptions = {},
): Promise<OfficeArchiveReader> {
  // zip.js 2.8.34 仅在该选项为 false 时才会先初始化当前上下文中的 WASM 解压流。
  const useCompressionStream = supportsNativeZipCore();
  const zipModule = useCompressionStream
    ? await import('@zip.js/zip.js/lib/zip-core-native.js')
    : await import('@zip.js/zip.js/lib/zip-core.js');
  return createZipJsOfficeArchiveReader(
    file,
    zipModule,
    useCompressionStream,
    options.signal,
  );
}

/**
 * 读取 Office ZIP 包中的全部文件，并限制同时解压的条目数量以降低瞬时资源峰值。
 */
export async function loadOfficeEntries(
  file: OfficeZipInput,
  options: OfficeArchiveOpenOptions = {},
): Promise<OfficeEntryMap> {
  const reader = await openOfficeArchive(file, options);
  try {
    return await reader.materialize(options.signal);
  } finally {
    await reader.close();
  }
}

/** 读取并解析指定 OOXML 文本条目。 */
export function readXml(entries: OfficeEntryMap, path: string) {
  const value = entries.get(path);
  return typeof value === 'string' ? value : '';
}

/** 读取指定 OOXML 二进制条目。 */
export function readBinary(entries: OfficeEntryMap, path: string) {
  const value = entries.get(path);
  return value instanceof Uint8Array ? value : undefined;
}
