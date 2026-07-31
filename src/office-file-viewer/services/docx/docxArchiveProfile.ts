import { openOfficeArchive } from '../../shared/ooxml/archive';
import type {
  OfficeArchiveEntry,
  OfficeArchiveReader,
} from '../../shared/ooxml/OfficeArchiveReader';
import { OFFICE_LARGE_FILE_THRESHOLDS } from '../performance/officePerformanceThresholds';

/** DOCX 压缩包的大文件判定指标。 */
export type DocxArchiveProfile = {
  /** 当前数据源或渲染器采用的工作模式。 */
  mode: 'materialized' | 'lazy';
  /** 压缩包中相关内容的压缩大小，单位为字节。 */
  compressedSize: number;
  /** 相关内容解压后的大小，单位为字节。 */
  uncompressedSize: number;
  /** DOCX 主文档 XML 的大小，单位为字节。 */
  mainDocumentSize: number;
  /** 压缩包内最大媒体文件的大小，单位为字节。 */
  largestMediaSize: number;
};

/** 附带性能档案的 DOCX 压缩包读取器。 */
export type ProfiledDocxArchive = {
  /** 用于按需读取源数据的读取器。 */
  reader: OfficeArchiveReader;
  /** 控制解析或渲染策略的性能档案。 */
  profile: DocxArchiveProfile;
};

/** 根据 ZIP 中央目录元数据计算 DOCX 物化或懒读取模式。 */
export function createDocxArchiveProfile(
  compressedSize: number,
  entries: readonly OfficeArchiveEntry[],
): DocxArchiveProfile {
  const uncompressedSize = entries.reduce(
    (total, entry) => total + entry.uncompressedSize,
    0,
  );
  const mainDocumentSize =
    entries.find((entry) => entry.path === 'word/document.xml')
      ?.uncompressedSize ?? 0;
  const largestMediaSize = entries
    .filter((entry) => entry.path.startsWith('word/media/'))
    .reduce((largest, entry) => Math.max(largest, entry.uncompressedSize), 0);
  const mode =
    compressedSize >= OFFICE_LARGE_FILE_THRESHOLDS.ooxmlCompressedBytes ||
    uncompressedSize >= OFFICE_LARGE_FILE_THRESHOLDS.ooxmlUncompressedBytes ||
    mainDocumentSize >= OFFICE_LARGE_FILE_THRESHOLDS.ooxmlMainXmlBytes ||
    largestMediaSize >= OFFICE_LARGE_FILE_THRESHOLDS.ooxmlSingleMediaBytes
      ? 'lazy'
      : 'materialized';
  return {
    mode,
    compressedSize,
    uncompressedSize,
    mainDocumentSize,
    largestMediaSize,
  };
}

/** 只读取 ZIP 中央目录形成 DOCX 提前画像，不解压正文或媒体。 */
export async function profileDocxArchive(
  file: File,
  signal?: AbortSignal,
): Promise<ProfiledDocxArchive> {
  const reader = await openOfficeArchive(file, { signal });
  try {
    const entries = reader.list();
    return {
      reader,
      profile: createDocxArchiveProfile(file.size, entries),
    };
  } catch (error) {
    await reader.close();
    throw error;
  }
}
