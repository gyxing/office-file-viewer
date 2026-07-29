import { openOfficeArchive } from '../../shared/ooxml/archive';
import type {
  OfficeArchiveEntry,
  OfficeArchiveReader,
} from '../../shared/ooxml/OfficeArchiveReader';
import { OFFICE_LARGE_FILE_THRESHOLDS } from '../performance/officePerformanceThresholds';

export type DocxArchiveProfile = {
  mode: 'materialized' | 'lazy';
  compressedSize: number;
  uncompressedSize: number;
  mainDocumentSize: number;
  largestMediaSize: number;
};

export type ProfiledDocxArchive = {
  reader: OfficeArchiveReader;
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
