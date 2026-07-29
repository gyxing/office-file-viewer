import { openOfficeArchive } from '../../shared/ooxml/archive';
import { imageMimeType } from '../../shared/ooxml/media';
import type {
  OfficeArchiveEntry,
  OfficeArchiveReader,
} from '../../shared/ooxml/OfficeArchiveReader';
import {
  attr,
  childrenByLocalName,
  descendantByLocalName,
  parseXml,
} from '../../shared/ooxml/xml';
import {
  createPresentationPerformanceProfile,
  type PresentationPerformanceProfile,
} from '../presentation/presentationPerformance';
import type { OfficeResourceSource } from '../resource-store';
import {
  buildPptxPackageState,
  readPresentationLayouts,
  readPresentationSize,
  readTableStyles,
  readTheme,
  throwIfPptxParseAborted,
} from './parsePptx';
import type {
  PptxPackageContext,
  PptxSlideDescriptor,
} from './PptxPackageContext';

export type PptxArchiveProfile = {
  performance: PresentationPerformanceProfile;
  compressedSize: number;
  uncompressedSize: number;
  slideCount: number;
  slideXmlBytes: number;
  mediaBytes: number;
  largestMediaSize: number;
};

export type ProfiledPptxArchive = {
  reader: OfficeArchiveReader;
  profile: PptxArchiveProfile;
};

/** 仅根据 ZIP 中央目录计算 PPTX 的提前性能画像。 */
export function createPptxArchiveProfile(
  compressedSize: number,
  entries: readonly OfficeArchiveEntry[],
): PptxArchiveProfile {
  const slideEntries = entries.filter((entry) =>
    /^ppt\/slides\/slide\d+\.xml$/i.test(entry.path),
  );
  const mediaEntries = entries.filter((entry) =>
    entry.path.startsWith('ppt/media/'),
  );
  const uncompressedSize = entries.reduce(
    (total, entry) => total + entry.uncompressedSize,
    0,
  );
  const slideXmlBytes = slideEntries.reduce(
    (total, entry) => total + entry.uncompressedSize,
    0,
  );
  const mediaBytes = mediaEntries.reduce(
    (total, entry) => total + entry.uncompressedSize,
    0,
  );
  const largestMediaSize = mediaEntries.reduce(
    (largest, entry) => Math.max(largest, entry.uncompressedSize),
    0,
  );
  return {
    performance: createPresentationPerformanceProfile({
      slideCount: slideEntries.length,
      compressedBytes: compressedSize,
      uncompressedBytes: uncompressedSize,
      mainXmlBytes: slideXmlBytes,
      mediaBytes,
      singleMediaBytes: largestMediaSize,
    }),
    compressedSize,
    uncompressedSize,
    slideCount: slideEntries.length,
    slideXmlBytes,
    mediaBytes,
    largestMediaSize,
  };
}

/** 打开 PPTX 并只读取 ZIP 中央目录，调用方负责转移或关闭 Reader。 */
export async function profilePptxArchive(
  file: File,
  signal?: AbortSignal,
): Promise<ProfiledPptxArchive> {
  const reader = await openOfficeArchive(file, { signal });
  try {
    return {
      reader,
      profile: createPptxArchiveProfile(file.size, reader.list()),
    };
  } catch (error) {
    await reader.close();
    throw error;
  }
}

async function readTextEntries(
  reader: OfficeArchiveReader,
  paths: readonly string[],
  signal?: AbortSignal,
) {
  const entries = new Map<string, string | Uint8Array>();
  const queue = [...new Set(paths)].filter((path) => reader.has(path));
  const workers = Array.from(
    { length: Math.min(4, Math.max(1, queue.length)) },
    async () => {
      while (queue.length) {
        throwIfPptxParseAborted(signal);
        const path = queue.shift();
        if (!path) return;
        entries.set(path, await reader.readText(path, signal));
      }
    },
  );
  await Promise.all(workers);
  return entries;
}

/** 读取主题、母版、版式、关系和 Slide 描述符，但不进入任何 Slide 正文。 */
export async function readPptxStructure(
  reader: OfficeArchiveReader,
  sessionId: string,
  profile: PptxArchiveProfile,
  signal?: AbortSignal,
): Promise<PptxPackageContext> {
  throwIfPptxParseAborted(signal);
  const archiveEntries = reader.list();
  const structuralPaths = archiveEntries
    .map((entry) => entry.path)
    .filter(
      (path) =>
        path === 'ppt/presentation.xml' ||
        path === 'ppt/theme/theme1.xml' ||
        path === 'ppt/tableStyles.xml' ||
        path.endsWith('.rels') ||
        /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(path) ||
        /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(path),
    );
  const entries = await readTextEntries(reader, structuralPaths, signal);
  const mediaByPath: Record<string, OfficeResourceSource> = {};
  const mediaByName: Record<string, OfficeResourceSource> = {};
  archiveEntries
    .filter((entry) => entry.path.startsWith('ppt/media/'))
    .forEach((entry) => {
      const source: OfficeResourceSource = {
        kind: 'lazy',
        id: `${sessionId}:pptx:${entry.path}`,
        mimeType: imageMimeType(entry.path),
        size: entry.uncompressedSize,
        load: (resourceSignal) =>
          reader.readBlob(
            entry.path,
            imageMimeType(entry.path),
            resourceSignal,
          ),
      };
      mediaByPath[entry.path] = source;
      mediaByName[entry.path.split('/').pop() ?? entry.path] = source;
    });
  const packageState = buildPptxPackageState(entries, {
    mediaByName,
    mediaByPath,
  });
  const presentationXml =
    (entries.get('ppt/presentation.xml') as string | undefined) ?? '';
  const presentationDocument = parseXml(presentationXml);
  const themeXml =
    (entries.get('ppt/theme/theme1.xml') as string | undefined) ?? '';
  const theme = themeXml
    ? readTheme(themeXml)
    : { colorScheme: {}, fontScheme: {}, colorMap: {} };
  const tableStyles = readTableStyles(
    (entries.get('ppt/tableStyles.xml') as string | undefined) ?? '',
    theme,
  );
  const size = presentationXml
    ? readPresentationSize(presentationXml)
    : { width: 960, height: 540 };
  const presentationRels =
    packageState.relationships['ppt/_rels/presentation.xml.rels'] ?? {};
  const slideIds = childrenByLocalName(
    descendantByLocalName(presentationDocument.documentElement, 'sldIdLst'),
    'sldId',
  );
  const descriptors: PptxSlideDescriptor[] = slideIds.map((node, index) => {
    const relId = attr(node, 'r:id');
    const target = relId ? presentationRels[relId] : undefined;
    const slidePath = target ?? `ppt/slides/slide${index + 1}.xml`;
    const relsPath = slidePath
      .replace(/^ppt\/slides\//, 'ppt/slides/_rels/')
      .concat('.rels');
    const notesPath = Object.values(
      packageState.relationships[relsPath] ?? {},
    ).find((value) => value.includes('notesSlides/'));
    const slideSize =
      archiveEntries.find((entry) => entry.path === slidePath)
        ?.uncompressedSize ?? 0;
    return {
      id: `slide-${index + 1}`,
      index: index + 1,
      hidden: attr(node, 'show') === '0',
      hasSpeakerNotes: Boolean(notesPath),
      estimatedElementCount: Math.max(1, Math.ceil(slideSize / 1024)),
      revision: 1,
      status: 'estimated',
      slidePath,
      relsPath,
      notesPath,
    };
  });
  const { masterDefinitions, masterLayoutDefinitions } =
    readPresentationLayouts(entries, packageState, theme, tableStyles);

  return {
    sessionId,
    reader,
    packageState,
    width: size.width,
    height: size.height,
    theme,
    tableStyles,
    masterDefinitions,
    layoutDefinitions: Object.values(masterLayoutDefinitions).flat(),
    descriptors,
  };
}
