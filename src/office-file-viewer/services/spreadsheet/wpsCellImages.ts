import { imageMimeType } from '../../shared/ooxml/media';
import type { OfficeArchiveReader } from '../../shared/ooxml/OfficeArchiveReader';
import { readRelationships } from '../../shared/ooxml/relationships';
import {
  attr,
  descendantByLocalName,
  descendantsByLocalName,
  parseXml,
} from '../../shared/ooxml/xml';
import { createSpreadsheetAxisIndex } from './SpreadsheetAxisIndex';
import type { SpreadsheetSheetLayout } from './SpreadsheetSource';
import type {
  SpreadsheetCell,
  SpreadsheetImage,
  SpreadsheetMerge,
} from './types';

/** WPS DISPIMG 公式引用的单元格图片位置。 */
export type WpsCellImagePlacement = {
  /** 图片标识。 */
  imageId: string;
  /** 当前处理的表格行。 */
  row: number;
  /** 当前处理的列定义。 */
  column: number;
};

/** WPS cellImages 清单中的图片资源定义。 */
export type WpsCellImageDefinition = {
  /** 图片标识。 */
  imageId: string;
  /** 面向用户展示的名称。 */
  name?: string;
  /** 图片无法显示时使用的替代文本。 */
  alt?: string;
  /** 图片路径。 */
  imagePath: string;
};

/** WPS 单元格图片的资源和锚点信息。 */
type WpsCellImageLayout = Pick<
  SpreadsheetSheetLayout,
  | 'rowCount'
  | 'columnCount'
  | 'defaultRowHeight'
  | 'defaultColumnWidth'
  | 'rows'
  | 'columns'
>;

/** 从 WPS DISPIMG 公式中提取图片标识的正则表达式。 */
const DISPIMG_ID_PATTERN = /\bDISPIMG\b[^"']*["']([^"']+)["']/i;

/** 从 XLS/XLSX 的 DISPIMG 公式或缓存值中读取图片标识。 */
export function readWpsCellImageId(
  ...values: Array<string | undefined>
): string | undefined {
  for (const value of values) {
    const imageId = value?.match(DISPIMG_ID_PATTERN)?.[1]?.trim();
    if (imageId) return imageId;
  }
  return undefined;
}

/** 识别单元格图片公式并返回其工作表位置。 */
export function readWpsCellImagePlacement(
  cell: Pick<
    SpreadsheetCell,
    'rowIndex' | 'columnIndex' | 'formula' | 'formulaTokens' | 'value'
  >,
): WpsCellImagePlacement | undefined {
  const imageId = readWpsCellImageId(
    cell.formula,
    cell.formulaTokens,
    cell.value,
  );
  return imageId
    ? { imageId, row: cell.rowIndex, column: cell.columnIndex }
    : undefined;
}

function findArchivePath(reader: OfficeArchiveReader, expectedPath: string) {
  const normalized = expectedPath.toLowerCase();
  return reader.list().find((entry) => entry.path.toLowerCase() === normalized)
    ?.path;
}

function relationshipPath(partPath: string) {
  const parts = partPath.split('/');
  const fileName = parts.pop() ?? partPath;
  return `${parts.join('/')}/_rels/${fileName}.rels`;
}

/** 解析 WPS cellImages 清单并将关系目标转换为可读取的包内路径。 */
export function parseWpsCellImageDefinitions(
  partXml: string,
  relsXml: string,
  relsPath: string,
  resolvePath: (target: string) => string | undefined,
) {
  const relationships = readRelationships(relsXml, relsPath);
  const document = parseXml(partXml);
  const definitions = new Map<string, WpsCellImageDefinition>();
  descendantsByLocalName(document.documentElement, 'cellImage').forEach(
    (node) => {
      const metadata = descendantByLocalName(node, 'cNvPr');
      const imageId = attr(metadata, 'name')?.trim();
      const blip = descendantByLocalName(node, 'blip');
      const relationshipId = attr(blip, 'r:embed') ?? attr(blip, 'embed');
      const target = relationshipId
        ? relationships[relationshipId]?.target
        : undefined;
      const imagePath = target ? resolvePath(target) : undefined;
      if (!imageId || !imagePath) return;
      definitions.set(imageId, {
        imageId,
        name: attr(metadata, 'descr') ?? imageId,
        alt: attr(metadata, 'descr'),
        imagePath,
      });
    },
  );
  return definitions;
}

async function readWpsCellImageDefinitions(
  reader: OfficeArchiveReader,
  signal?: AbortSignal,
) {
  const partPath = reader
    .list()
    .find((entry) => /(^|\/)cellimages\.xml$/i.test(entry.path))?.path;
  if (!partPath) return new Map<string, WpsCellImageDefinition>();
  const relsPath = findArchivePath(reader, relationshipPath(partPath));
  if (!relsPath) return new Map<string, WpsCellImageDefinition>();
  const [partXml, relsXml] = await Promise.all([
    reader.readText(partPath, signal),
    reader.readText(relsPath, signal),
  ]);
  return parseWpsCellImageDefinitions(partXml, relsXml, relsPath, (target) =>
    findArchivePath(reader, target),
  );
}

function findPlacementMerge(
  placement: WpsCellImagePlacement,
  merges: readonly SpreadsheetMerge[],
) {
  return merges.find(
    (merge) =>
      placement.row >= merge.startRow &&
      placement.row <= merge.endRow &&
      placement.column >= merge.startColumn &&
      placement.column <= merge.endColumn,
  );
}

/**
 * 将 WPS cellImages 包中的资源按 DISPIMG 单元格及其合并区域转换为标准图片。
 */
export async function loadWpsCellImages(
  reader: OfficeArchiveReader,
  sessionId: string,
  placements: readonly WpsCellImagePlacement[],
  layout: WpsCellImageLayout,
  merges: readonly SpreadsheetMerge[],
  signal?: AbortSignal,
): Promise<SpreadsheetImage[]> {
  if (!placements.length) return [];
  const definitions = await readWpsCellImageDefinitions(reader, signal);
  if (!definitions.size) return [];
  const entries = new Map(reader.list().map((entry) => [entry.path, entry]));
  const rowAxis = createSpreadsheetAxisIndex(
    layout.rowCount,
    layout.defaultRowHeight,
    layout.rows,
  );
  const columnAxis = createSpreadsheetAxisIndex(
    layout.columnCount,
    layout.defaultColumnWidth,
    layout.columns,
  );
  return placements.flatMap((placement, index) => {
    const definition = definitions.get(placement.imageId);
    if (!definition) return [];
    const merge = findPlacementMerge(placement, merges);
    const startRow = merge?.startRow ?? placement.row;
    const startColumn = merge?.startColumn ?? placement.column;
    const endRow = merge?.endRow ?? placement.row;
    const endColumn = merge?.endColumn ?? placement.column;
    const x = columnAxis.offsetAt(startColumn);
    const y = rowAxis.offsetAt(startRow);
    const right = columnAxis.offsetAt(endColumn + 1);
    const bottom = rowAxis.offsetAt(endRow + 1);
    const entry = entries.get(definition.imagePath);
    const mimeType = imageMimeType(definition.imagePath);
    return [
      {
        id: `${sessionId}:wps-cell-image:${placement.imageId}:${placement.row}:${placement.column}:${index}`,
        name: definition.name,
        alt: definition.alt,
        src: {
          kind: 'lazy' as const,
          id: `${sessionId}:wps-cell-image-resource:${definition.imagePath}`,
          mimeType,
          size: entry?.uncompressedSize ?? 0,
          load: (resourceSignal?: AbortSignal) =>
            reader.readBlob(definition.imagePath, mimeType, resourceSignal),
        },
        from: {
          row: startRow,
          column: startColumn,
          rowOffset: 0,
          columnOffset: 0,
        },
        to: {
          row: endRow + 1,
          column: endColumn + 1,
          rowOffset: 0,
          columnOffset: 0,
        },
        x,
        y,
        width: Math.max(1, right - x),
        height: Math.max(1, bottom - y),
      },
    ];
  });
}
