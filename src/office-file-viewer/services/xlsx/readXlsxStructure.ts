import { openOfficeArchive } from '../../shared/ooxml/archive';
import type {
  OfficeArchiveEntry,
  OfficeArchiveReader,
} from '../../shared/ooxml/OfficeArchiveReader';
import { readOfficeXmlEvents } from '../../shared/ooxml/OfficeXmlEventReader';
import { readRelationships } from '../../shared/ooxml/relationships';
import { readOfficeTheme } from '../../shared/ooxml/theme';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  descendantsByLocalName,
  parseXml,
} from '../../shared/ooxml/xml';
import type { OfficeArchiveResourcePolicy } from '../../shared/resource/OfficeResourcePolicy';
import {
  createSpreadsheetPerformanceProfile,
  type SpreadsheetPerformanceProfile,
} from '../spreadsheet/spreadsheetPerformance';
import type { SpreadsheetRange } from '../spreadsheet/types';
import { decodeMojibake, parseRange, parseStyles } from './xlsxCellFormatting';
import type {
  XlsxPackageContext,
  XlsxSheetDescriptor,
} from './XlsxPackageContext';
import { createXlsxSharedStringSource } from './XlsxSharedStringSource';

/** ZIP 中央目录和 Sheet 估算得到的 XLSX 性能画像。 */
export type XlsxArchiveProfile = {
  /** 压缩包中相关内容的压缩大小，单位为字节。 */
  compressedSize: number;
  /** 相关内容解压后的大小，单位为字节。 */
  uncompressedSize: number;
  /** 按工作表路径索引的性能档案。 */
  profiles: ReadonlyMap<string, SpreadsheetPerformanceProfile>;
  /** 是否至少有一张工作表需要启用按需数据源。 */
  requiresSource: boolean;
};

/** 已打开且尚未转移所有权的 XLSX Reader。 */
export type ProfiledXlsxArchive = {
  /** 用于按需读取源数据的读取器。 */
  reader: OfficeArchiveReader;
  /** 压缩包或复合文档包含的条目。 */
  entries: readonly OfficeArchiveEntry[];
  /** 压缩包中相关内容的压缩大小，单位为字节。 */
  compressedSize: number;
  /** 相关内容解压后的大小，单位为字节。 */
  uncompressedSize: number;
};

function sheetRelationshipPath(path: string) {
  const parts = path.split('/');
  const fileName = parts.pop() ?? path;
  return `${parts.join('/')}/_rels/${fileName}.rels`;
}

async function readSheetEstimatedRange(
  reader: OfficeArchiveReader,
  path: string,
  signal?: AbortSignal,
): Promise<SpreadsheetRange | undefined> {
  if (!reader.has(path)) return undefined;
  const stream = await reader.openStream(path, signal);
  for await (const event of readOfficeXmlEvents(stream, signal)) {
    if (event.type !== 'open') continue;
    if (event.localName === 'dimension') {
      return parseRange(event.attributes.get('ref'));
    }
    if (event.localName === 'sheetData') return undefined;
  }
  return undefined;
}

async function readTextEntries(
  reader: OfficeArchiveReader,
  paths: readonly string[],
  signal?: AbortSignal,
) {
  const values = new Map<string, string>();
  const queue = [...new Set(paths)].filter((path) => reader.has(path));
  await Promise.all(
    Array.from({ length: Math.min(4, Math.max(1, queue.length)) }, async () => {
      while (queue.length) {
        const path = queue.shift();
        if (!path) return;
        values.set(path, await reader.readText(path, signal));
      }
    }),
  );
  return values;
}

/** 打开 XLSX 并只读取 ZIP 中央目录。 */
export async function profileXlsxArchive(
  file: File,
  signal?: AbortSignal,
  resourcePolicy?: OfficeArchiveResourcePolicy,
): Promise<ProfiledXlsxArchive> {
  const reader = await openOfficeArchive(file, { signal, resourcePolicy });
  try {
    const entries = reader.list();
    return {
      reader,
      entries,
      compressedSize: file.size,
      uncompressedSize: entries.reduce(
        (total, entry) => total + entry.uncompressedSize,
        0,
      ),
    };
  } catch (error) {
    await reader.close();
    throw error;
  }
}

/** 读取工作簿、样式、主题和 Sheet 描述符，不进入任何 Sheet 正文。 */
export async function readXlsxStructure(
  archive: ProfiledXlsxArchive,
  sessionId: string,
  signal?: AbortSignal,
): Promise<{ context: XlsxPackageContext; profile: XlsxArchiveProfile }> {
  const { reader, entries } = archive;
  const relationshipPaths = entries
    .map((entry) => entry.path)
    .filter((path) => path.endsWith('.rels'));
  const values = await readTextEntries(
    reader,
    [
      'xl/workbook.xml',
      'xl/styles.xml',
      'xl/theme/theme1.xml',
      ...relationshipPaths,
    ],
    signal,
  );
  const relationships: XlsxPackageContext['relationships'] = {};
  relationshipPaths.forEach((path) => {
    const xml = values.get(path);
    if (xml) relationships[path] = readRelationships(xml, path);
  });
  const workbookXml = values.get('xl/workbook.xml') ?? '';
  const workbook = parseXml(workbookXml);
  const definedNames = Object.fromEntries(
    descendantsByLocalName(workbook.documentElement, 'definedName')
      .map((node) => [attr(node, 'name'), node.textContent?.trim()] as const)
      .filter(
        (entry): entry is readonly [string, string] =>
          Boolean(entry[0] && entry[1]),
      ),
  );
  const workbookRels = relationships['xl/_rels/workbook.xml.rels'] ?? {};
  const sheetEntries = childrenByLocalName(
    childByLocalName(workbook.documentElement, 'sheets'),
    'sheet',
  )
    .map((node, sourceIndex) => ({ node, sourceIndex }))
    .filter(({ node }) => {
      // 隐藏工作表属于工作簿内部数据，不应出现在预览标签栏中。
      const state = attr(node, 'state');
      return state !== 'hidden' && state !== 'veryHidden';
    });
  const entryByPath = new Map(entries.map((entry) => [entry.path, entry]));
  const dimensions = await Promise.all(
    sheetEntries.map(async ({ node, sourceIndex }) => {
      const relationshipId = attr(node, 'r:id') ?? attr(node, 'id') ?? '';
      const path =
        workbookRels[relationshipId]?.target ??
        `xl/worksheets/sheet${sourceIndex + 1}.xml`;
      return readSheetEstimatedRange(reader, path, signal);
    }),
  );
  const profiles = new Map<string, SpreadsheetPerformanceProfile>();
  const descriptors: XlsxSheetDescriptor[] = sheetEntries.map(
    ({ node, sourceIndex }, index) => {
      const relationshipId = attr(node, 'r:id') ?? attr(node, 'id') ?? '';
      const relationship = workbookRels[relationshipId];
      const path =
        relationship?.target ?? `xl/worksheets/sheet${sourceIndex + 1}.xml`;
      const id = attr(node, 'sheetId') ?? String(sourceIndex + 1);
      const estimatedRange = dimensions[index];
      const entry = entryByPath.get(path);
      const performance = createSpreadsheetPerformanceProfile({
        rowCount: estimatedRange?.endRow ?? 1,
        columnCount: estimatedRange?.endColumn ?? 1,
        compressedBytes: archive.compressedSize,
        uncompressedBytes: archive.uncompressedSize,
        sheetBytes: entry?.uncompressedSize ?? 0,
      });
      profiles.set(id, performance);
      return {
        id,
        name: decodeMojibake(attr(node, 'name') ?? `Sheet ${sourceIndex + 1}`),
        path,
        kind: path.includes('/chartsheets/') ? 'chart' : 'worksheet',
        rowCount: estimatedRange?.endRow ?? 1,
        columnCount: estimatedRange?.endColumn ?? 1,
        revision: 1,
        status: 'estimated',
        sheetBytes: entry?.uncompressedSize ?? 0,
        relsPath: sheetRelationshipPath(path),
        performance,
      };
    },
  );
  const theme = readOfficeTheme(values.get('xl/theme/theme1.xml'));
  const context: XlsxPackageContext = {
    sessionId,
    reader,
    relationships,
    styles: parseStyles(values.get('xl/styles.xml') ?? '', theme),
    theme,
    sharedStrings: createXlsxSharedStringSource(reader),
    descriptors,
    definedNames,
  };
  return {
    context,
    profile: {
      compressedSize: archive.compressedSize,
      uncompressedSize: archive.uncompressedSize,
      profiles,
      requiresSource:
        descriptors.some(
          (descriptor) =>
            descriptor.performance.sheetMode === 'lazy' ||
            descriptor.performance.gridMode !== 'table',
        ) ||
        // WPS 单元格图片保持为按需资源，必须让 Source 持有 ZIP Reader 生命周期。
        entries.some((entry) => /(^|\/)cellimages\.xml$/i.test(entry.path)) ||
        // EMF/WMF 转换成本高且依赖 Canvas，只在图片进入可见范围时处理。
        entries.some((entry) =>
          /(^|\/)xl\/media\/.*\.(emf|wmf)$/i.test(entry.path),
        ),
    },
  };
}
