import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import { readXml } from '../../shared/ooxml/archive';
import { parseOfficeChartXml } from '../../shared/ooxml/charts';
import {
  resolvePackageMediaRef,
  type OfficeRelationship,
} from '../../shared/ooxml/media';
import type { OfficeTheme } from '../../shared/ooxml/theme';
import { emuToPx } from '../../shared/ooxml/units';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  descendantByLocalName,
  descendantsByLocalName,
  parseXml,
  textContent,
} from '../../shared/ooxml/xml';
import {
  applyConditionalFormatting,
  applySpreadsheetTableSemantics,
} from '../spreadsheet/semantics';
import {
  parseWpsCellImageDefinitions,
  readWpsCellImageId,
  type WpsCellImagePlacement,
} from '../spreadsheet/wpsCellImages';
import { mergeXlsxPreviewImages } from './loadXlsxOlePreviewImages';
import { parseMaterializedXlsxComments } from './parseXlsxComments';
import { parseMaterializedXlsxConditionalFormatting } from './parseXlsxConditionalFormatting';
import {
  applyStaticXlsxFormulaHyperlink,
  applyXlsxHyperlinkRanges,
  parseXlsxDrawingHyperlink,
  parseXlsxHyperlink,
} from './parseXlsxHyperlinks';
import { parseMaterializedXlsxPane } from './parseXlsxPane';
import { parseMaterializedXlsxTables } from './parseXlsxTables';
import type {
  XlsxCell,
  XlsxChart,
  XlsxColumn,
  XlsxImage,
  XlsxMerge,
  XlsxRow,
  XlsxSheet,
} from './types';
import {
  DEFAULT_COLUMN_WIDTH,
  DEFAULT_ROW_HEIGHT_POINTS,
  columnIndexToLabel,
  decodeMojibake,
  excelWidthToPx,
  parseCellRef,
  parseRange,
  pointToPx,
  resolveStyle,
  resolveXlsxMaxDigitWidth,
  type StyleBook,
} from './xlsxCellFormatting';

/** 完整物化 XLSX 时共享的包资源与关系状态。 */
export type MaterializedXlsxPackageState = {
  /** 按包内路径索引的 OOXML 条目。 */
  entries: OfficeEntryMap;
  /** 按关系文件路径组织的 OOXML 关系映射。 */
  relationships: Record<string, Record<string, OfficeRelationship>>;
  /** 按 OOXML 包内路径索引的媒体资源映射。 */
  mediaByPath: Record<string, string>;
  /** 按媒体文件名索引的资源映射。 */
  mediaByName: Record<string, string>;
  /** 当前文档使用的主题颜色和字体配置。 */
  theme: OfficeTheme;
};

/** 解析工作表尺寸时允许补充渲染的最大空白行数。 */
const MAX_RENDERED_EMPTY_ROWS = 200;

/** 解析工作表尺寸时允许补充渲染的最大空白列数。 */
const MAX_RENDERED_EMPTY_COLUMNS = 80;

/** 还原 XLSX 行高和列宽所需的默认度量。 */
type SheetMetrics = {
  /** 工作表默认列宽，单位为标准化渲染像素。 */
  defaultColumnWidth: number;
  /** 工作表默认行高，单位为标准化渲染像素。 */
  defaultRowHeight: number;
  /** Normal 字体中数字 0 的最大像素宽度，用于还原 OOXML 字符列宽。 */
  maxDigitWidth: number;
};

function readPlainText(node: Element | null | undefined) {
  if (!node) return '';
  return decodeMojibake(
    descendantsByLocalName(node, 't')
      .map((item) => textContent(item))
      .join(''),
  );
}

/** 读取完整物化路径使用的共享字符串表。 */
export function readMaterializedXlsxSharedStrings(xml: string) {
  if (!xml) return [];
  const doc = parseXml(xml);
  return childrenByLocalName(doc.documentElement, 'si').map(readPlainText);
}

/** 返回目标列宽，未显式设置时使用工作表默认列宽。 */
function getColumnWidth(
  columns: XlsxColumn[],
  columnIndex: number,
  metrics: SheetMetrics,
) {
  return columns[columnIndex - 1]?.width ?? metrics.defaultColumnWidth;
}

/** 返回目标行高，未显式设置时使用工作表默认行高。 */
function getRowHeight(
  rowHeights: Map<number, number>,
  rowIndex: number,
  metrics: SheetMetrics,
) {
  return rowHeights.get(rowIndex) ?? metrics.defaultRowHeight;
}

function anchorPosition(
  anchor: {
    /** 锚点单元格的零基行索引。 */
    row: number;
    /** 锚点单元格的零基列索引。 */
    column: number;
    /** 锚点在单元格内的纵向偏移。 */
    rowOffset: number;
    /** 锚点在单元格内的横向偏移。 */
    columnOffset: number;
  },
  columns: XlsxColumn[],
  rowHeights: Map<number, number>,
  metrics: SheetMetrics,
) {
  let x = 0;
  for (let column = 1; column < anchor.column; column += 1) {
    x += getColumnWidth(columns, column, metrics);
  }

  let y = 0;
  for (let row = 1; row < anchor.row; row += 1) {
    y += getRowHeight(rowHeights, row, metrics);
  }

  return {
    x: x + anchor.columnOffset,
    y: y + anchor.rowOffset,
  };
}

function readSheetMetrics(
  sheetNode: Element,
  styleBook: StyleBook,
): SheetMetrics {
  const sheetFormat = childByLocalName(sheetNode, 'sheetFormatPr');
  const maxDigitWidth = resolveXlsxMaxDigitWidth(styleBook.fonts[0]);
  const defaultColumnWidth = attr(sheetFormat, 'defaultColWidth');
  return {
    defaultColumnWidth: defaultColumnWidth
      ? excelWidthToPx(
          Number(defaultColumnWidth),
          DEFAULT_COLUMN_WIDTH,
          maxDigitWidth,
        )
      : DEFAULT_COLUMN_WIDTH,
    defaultRowHeight: pointToPx(
      Number(
        attr(sheetFormat, 'defaultRowHeight') ?? DEFAULT_ROW_HEIGHT_POINTS,
      ),
    ),
    maxDigitWidth,
  };
}

function readColumns(
  sheetNode: Element,
  maxColumn: number,
  metrics: SheetMetrics,
): XlsxColumn[] {
  const widths = new Map<number, XlsxColumn>();
  descendantsByLocalName(sheetNode, 'col').forEach((node) => {
    const min = Number(attr(node, 'min') ?? 1);
    const max = Math.min(
      Number(attr(node, 'max') ?? min),
      Math.max(maxColumn, MAX_RENDERED_EMPTY_COLUMNS),
    );
    for (let index = min; index <= max; index += 1) {
      widths.set(index, {
        index,
        label: columnIndexToLabel(index),
        width: excelWidthToPx(
          Number(attr(node, 'width')),
          metrics.defaultColumnWidth,
          metrics.maxDigitWidth,
        ),
        hidden: attr(node, 'hidden') === '1',
      });
    }
  });

  return Array.from({ length: maxColumn }, (_, itemIndex) => {
    const index = itemIndex + 1;
    return (
      widths.get(index) ?? {
        index,
        label: columnIndexToLabel(index),
        width: metrics.defaultColumnWidth,
      }
    );
  });
}

function readAnchorPoint(node: Element | null) {
  return {
    column: Number(textContent(childByLocalName(node, 'col'))) + 1,
    // 标准模型的锚点偏移统一使用 CSS 像素，避免渲染阶段把 EMU 当像素再次参与比例换算。
    columnOffset: emuToPx(
      Number(textContent(childByLocalName(node, 'colOff')) || 0),
    ),
    row: Number(textContent(childByLocalName(node, 'row'))) + 1,
    rowOffset: emuToPx(
      Number(textContent(childByLocalName(node, 'rowOff')) || 0),
    ),
  };
}

/** 按源顺序读取 XLSX 支持的双单元格和单单元格绘图锚点。 */
function readDrawingAnchorNodes(root: Element) {
  return Array.from(root.children).filter(
    (node) =>
      node.localName === 'twoCellAnchor' || node.localName === 'oneCellAnchor',
  );
}

function readDrawingExtent(anchorNode: Element) {
  const extent = childByLocalName(anchorNode, 'ext');
  if (!extent) return undefined;
  return {
    width: emuToPx(Number(attr(extent, 'cx')) || 0),
    height: emuToPx(Number(attr(extent, 'cy')) || 0),
  };
}

/** 将绝对像素位置重新映射为单元格锚点，供 oneCellAnchor 补齐终点。 */
function anchorPointAtPosition(
  x: number,
  y: number,
  columns: XlsxColumn[],
  rowHeights: Map<number, number>,
  metrics: SheetMetrics,
) {
  let column = 1;
  let columnStart = 0;
  while (column < MAX_RENDERED_EMPTY_COLUMNS) {
    const width = getColumnWidth(columns, column, metrics);
    if (columnStart + width > x) break;
    columnStart += width;
    column += 1;
  }
  let row = 1;
  let rowStart = 0;
  while (row < MAX_RENDERED_EMPTY_ROWS) {
    const height = getRowHeight(rowHeights, row, metrics);
    if (rowStart + height > y) break;
    rowStart += height;
    row += 1;
  }
  return {
    column,
    columnOffset: Math.max(0, x - columnStart),
    row,
    rowOffset: Math.max(0, y - rowStart),
  };
}

function resolveDrawingAnchorGeometry(
  anchorNode: Element,
  columns: XlsxColumn[],
  rowHeights: Map<number, number>,
  metrics: SheetMetrics,
) {
  const from = readAnchorPoint(childByLocalName(anchorNode, 'from'));
  const start = anchorPosition(from, columns, rowHeights, metrics);
  const toNode = childByLocalName(anchorNode, 'to');
  if (toNode) {
    const to = readAnchorPoint(toNode);
    return {
      from,
      to,
      start,
      end: anchorPosition(to, columns, rowHeights, metrics),
    };
  }
  const extent = readDrawingExtent(anchorNode);
  if (!extent) return undefined;
  const end = { x: start.x + extent.width, y: start.y + extent.height };
  return {
    from,
    to: anchorPointAtPosition(end.x, end.y, columns, rowHeights, metrics),
    start,
    end,
  };
}

/** 解析并确定 `resolveMediaRef` 对应的引用或配置。 */
function resolveMediaRef(
  target: string | undefined,
  packageState: MaterializedXlsxPackageState,
) {
  return resolvePackageMediaRef(
    target,
    packageState.mediaByPath,
    packageState.mediaByName,
    'xl',
  );
}

/** 解析并确定 `resolveXmlTarget` 对应的引用或配置。 */
function resolveXmlTarget(
  target: string | undefined,
  packageState: MaterializedXlsxPackageState,
) {
  if (!target) return undefined;
  const normalized = target.replace(/^\.\.\//, '');
  return packageState.entries.get(normalized) ? normalized : target;
}

function readDrawingXml(
  sheetNode: Element,
  sheetPath: string,
  packageState: MaterializedXlsxPackageState,
) {
  const drawing = descendantByLocalName(sheetNode, 'drawing');
  const drawingRelId = attr(drawing, 'r:id') ?? attr(drawing, 'id');
  if (!drawingRelId) return undefined;

  const sheetRelPath = sheetPath
    .replace(/^xl\/worksheets\//, 'xl/worksheets/_rels/')
    .concat('.rels');
  const drawingPath =
    packageState.relationships[sheetRelPath]?.[drawingRelId]?.target;
  const drawingXml = drawingPath
    ? readXml(packageState.entries, drawingPath)
    : '';
  return drawingPath && drawingXml ? { drawingPath, drawingXml } : undefined;
}

function readDrawingBounds(
  sheetNode: Element,
  sheetPath: string,
  packageState: MaterializedXlsxPackageState,
  metrics: SheetMetrics,
) {
  const drawing = readDrawingXml(sheetNode, sheetPath, packageState);
  if (!drawing) return undefined;
  const drawingDoc = parseXml(drawing.drawingXml);
  let maxRow = 0;
  let maxColumn = 0;
  readDrawingAnchorNodes(drawingDoc.documentElement).forEach((anchorNode) => {
    const toNode = childByLocalName(anchorNode, 'to');
    if (toNode) {
      const to = readAnchorPoint(toNode);
      maxRow = Math.max(maxRow, to.row);
      maxColumn = Math.max(maxColumn, to.column);
      return;
    }
    const from = readAnchorPoint(childByLocalName(anchorNode, 'from'));
    const extent = readDrawingExtent(anchorNode);
    if (!extent) return;
    // 精确行列尺寸尚未建立，先按默认尺寸扩展边界，后续再用真实度量计算终点。
    maxRow = Math.max(
      maxRow,
      from.row +
        Math.ceil((from.rowOffset + extent.height) / metrics.defaultRowHeight),
    );
    maxColumn = Math.max(
      maxColumn,
      from.column +
        Math.ceil(
          (from.columnOffset + extent.width) / metrics.defaultColumnWidth,
        ),
    );
  });
  return maxRow || maxColumn ? { maxRow, maxColumn } : undefined;
}

/** 读取 worksheet/objectPr 预览图的最远锚点。 */
function readOlePreviewBounds(sheetNode: Element) {
  let maxRow = 0;
  let maxColumn = 0;
  descendantsByLocalName(sheetNode, 'objectPr').forEach((objectPr) => {
    const anchor = descendantByLocalName(objectPr, 'anchor');
    const to = readAnchorPoint(childByLocalName(anchor, 'to'));
    maxRow = Math.max(maxRow, to.row);
    maxColumn = Math.max(maxColumn, to.column);
  });
  return maxRow || maxColumn ? { maxRow, maxColumn } : undefined;
}

function readSheetCharts(
  sheetNode: Element,
  sheetPath: string,
  packageState: MaterializedXlsxPackageState,
  columns: XlsxColumn[],
  rowHeights: Map<number, number>,
  metrics: SheetMetrics,
) {
  const drawing = readDrawingXml(sheetNode, sheetPath, packageState);
  if (!drawing) return [];

  const drawingRelPath = drawing.drawingPath
    .replace(/^xl\/drawings\//, 'xl/drawings/_rels/')
    .concat('.rels');
  const drawingRels = packageState.relationships[drawingRelPath] ?? {};
  const drawingDoc = parseXml(drawing.drawingXml);

  return readDrawingAnchorNodes(drawingDoc.documentElement)
    .map((anchorNode, index): XlsxChart | undefined => {
      const graphicFrame = childByLocalName(anchorNode, 'graphicFrame');
      const chartNode = descendantByLocalName(graphicFrame, 'chart');
      const relId = attr(chartNode, 'r:id') ?? attr(chartNode, 'id');
      const target = relId ? drawingRels[relId]?.target : undefined;
      const chartPath = resolveXmlTarget(target, packageState);
      const xml = chartPath
        ? (packageState.entries.get(chartPath) as string | undefined)
        : undefined;
      if (!xml) return undefined;

      const geometry = resolveDrawingAnchorGeometry(
        anchorNode,
        columns,
        rowHeights,
        metrics,
      );
      if (!geometry) return undefined;
      const chart = parseOfficeChartXml(xml, packageState.theme);
      const name = attr(descendantByLocalName(anchorNode, 'cNvPr'), 'name');

      return {
        id: `${drawing.drawingPath}-chart-${index + 1}`,
        title: name,
        chart,
        from: geometry.from,
        to: geometry.to,
        x: geometry.start.x,
        y: geometry.start.y,
        width: Math.max(1, geometry.end.x - geometry.start.x),
        height: Math.max(1, geometry.end.y - geometry.start.y),
        hyperlink: parseXlsxDrawingHyperlink(anchorNode, drawingRels),
      };
    })
    .filter(Boolean) as XlsxChart[];
}

function readSheetImages(
  sheetNode: Element,
  sheetPath: string,
  packageState: MaterializedXlsxPackageState,
  columns: XlsxColumn[],
  rowHeights: Map<number, number>,
  metrics: SheetMetrics,
) {
  const drawing = readDrawingXml(sheetNode, sheetPath, packageState);
  if (!drawing) return [];

  const drawingRelPath = drawing.drawingPath
    .replace(/^xl\/drawings\//, 'xl/drawings/_rels/')
    .concat('.rels');
  const drawingRels = packageState.relationships[drawingRelPath] ?? {};
  const drawingDoc = parseXml(drawing.drawingXml);

  return readDrawingAnchorNodes(drawingDoc.documentElement)
    .map((anchorNode, index): XlsxImage | undefined => {
      const geometry = resolveDrawingAnchorGeometry(
        anchorNode,
        columns,
        rowHeights,
        metrics,
      );
      if (!geometry) return undefined;
      const blip = descendantByLocalName(anchorNode, 'blip');
      const embed = attr(blip, 'r:embed') ?? attr(blip, 'embed');
      const target = embed ? drawingRels[embed]?.target : undefined;
      const src = resolveMediaRef(target, packageState);
      if (!src) return undefined;

      const name = attr(descendantByLocalName(anchorNode, 'cNvPr'), 'name');

      return {
        id: `${drawing.drawingPath}-${index + 1}`,
        name,
        alt: name,
        src,
        from: geometry.from,
        to: geometry.to,
        x: geometry.start.x,
        y: geometry.start.y,
        width: Math.max(1, geometry.end.x - geometry.start.x),
        height: Math.max(1, geometry.end.y - geometry.start.y),
        hyperlink: parseXlsxDrawingHyperlink(anchorNode, drawingRels),
      };
    })
    .filter(Boolean) as XlsxImage[];
}

/** 读取嵌入对象 objectPr 记录的预览图及单元格锚点。 */
function readSheetOlePreviewImages(
  sheetNode: Element,
  sheetPath: string,
  packageState: MaterializedXlsxPackageState,
  columns: XlsxColumn[],
  rowHeights: Map<number, number>,
  metrics: SheetMetrics,
) {
  const sheetRelPath = sheetPath
    .replace(/^xl\/worksheets\//, 'xl/worksheets/_rels/')
    .concat('.rels');
  const relationships = packageState.relationships[sheetRelPath] ?? {};
  return descendantsByLocalName(sheetNode, 'objectPr')
    .map((objectPr, index): XlsxImage | undefined => {
      const relationshipId = attr(objectPr, 'r:id') ?? attr(objectPr, 'id');
      const target = relationshipId
        ? relationships[relationshipId]?.target
        : undefined;
      const src = resolveMediaRef(target, packageState);
      const anchor = descendantByLocalName(objectPr, 'anchor');
      if (!src || !anchor) return undefined;
      const from = readAnchorPoint(childByLocalName(anchor, 'from'));
      const to = readAnchorPoint(childByLocalName(anchor, 'to'));
      const start = anchorPosition(from, columns, rowHeights, metrics);
      const end = anchorPosition(to, columns, rowHeights, metrics);
      return {
        id: `${sheetPath}-ole-preview-${index + 1}`,
        name: '嵌入对象预览',
        alt: '嵌入对象预览',
        src,
        from,
        to,
        x: start.x,
        y: start.y,
        width: Math.max(1, end.x - start.x),
        height: Math.max(1, end.y - start.y),
      };
    })
    .filter(Boolean) as XlsxImage[];
}

function findPackagePath(entries: OfficeEntryMap, expectedPath: string) {
  const normalized = expectedPath.toLowerCase();
  return [...entries.keys()].find((path) => path.toLowerCase() === normalized);
}

function cellImageRelationshipPath(partPath: string) {
  const parts = partPath.split('/');
  const fileName = parts.pop() ?? partPath;
  return `${parts.join('/')}/_rels/${fileName}.rels`;
}

/** 读取 WPS DISPIMG 单元格图片，并按公式所在合并区域建立锚点。 */
function readSheetCellImages(
  placements: readonly WpsCellImagePlacement[],
  merges: readonly XlsxMerge[],
  packageState: MaterializedXlsxPackageState,
  columns: XlsxColumn[],
  rowHeights: Map<number, number>,
  metrics: SheetMetrics,
) {
  if (!placements.length) return [];
  const partPath = [...packageState.entries.keys()].find((path) =>
    /(^|\/)cellimages\.xml$/i.test(path),
  );
  if (!partPath) return [];
  const relsPath = findPackagePath(
    packageState.entries,
    cellImageRelationshipPath(partPath),
  );
  const partXml = packageState.entries.get(partPath);
  const relsXml = relsPath ? packageState.entries.get(relsPath) : undefined;
  if (typeof partXml !== 'string' || typeof relsXml !== 'string') return [];
  const definitions = parseWpsCellImageDefinitions(
    partXml,
    relsXml,
    relsPath!,
    (target) => findPackagePath(packageState.entries, target),
  );
  return placements.flatMap((placement, index): XlsxImage[] => {
    const definition = definitions.get(placement.imageId);
    const src = resolveMediaRef(definition?.imagePath, packageState);
    if (!definition || !src) return [];
    const merge = merges.find(
      (item) =>
        placement.row >= item.startRow &&
        placement.row <= item.endRow &&
        placement.column >= item.startColumn &&
        placement.column <= item.endColumn,
    );
    const startRow = merge?.startRow ?? placement.row;
    const startColumn = merge?.startColumn ?? placement.column;
    const endRow = merge?.endRow ?? placement.row;
    const endColumn = merge?.endColumn ?? placement.column;
    const from = {
      row: startRow,
      column: startColumn,
      rowOffset: 0,
      columnOffset: 0,
    };
    const to = {
      row: endRow + 1,
      column: endColumn + 1,
      rowOffset: 0,
      columnOffset: 0,
    };
    const start = anchorPosition(from, columns, rowHeights, metrics);
    const end = anchorPosition(to, columns, rowHeights, metrics);
    return [
      {
        id: `${partPath}-${placement.imageId}-${placement.row}-${placement.column}-${index}`,
        name: definition.name,
        alt: definition.alt,
        src,
        from,
        to,
        x: start.x,
        y: start.y,
        width: Math.max(1, end.x - start.x),
        height: Math.max(1, end.y - start.y),
      },
    ];
  });
}

function readCellValue(cellNode: Element, sharedStrings: string[]) {
  const type = attr(cellNode, 't');
  const valueNode = childByLocalName(cellNode, 'v');
  const rawValue = textContent(valueNode);

  if (type === 's') {
    return {
      rawValue,
      value: sharedStrings[Number(rawValue)] ?? '',
    };
  }

  if (type === 'inlineStr') {
    return {
      rawValue,
      value: readPlainText(childByLocalName(cellNode, 'is')),
    };
  }

  if (type === 'b') {
    return {
      rawValue,
      value: rawValue === '1' ? 'TRUE' : 'FALSE',
    };
  }

  return {
    rawValue,
    value: rawValue,
  };
}

function readMerges(sheetNode: Element) {
  const mergeCells = descendantByLocalName(sheetNode, 'mergeCells');
  return childrenByLocalName(mergeCells, 'mergeCell')
    .map((node) => {
      const ref = attr(node, 'ref') ?? '';
      const range = parseRange(ref);
      return range ? { ref, ...range } : undefined;
    })
    .filter(Boolean) as XlsxMerge[];
}

/** 标记合并区域根单元格的跨度，并隐藏其余占位单元格。 */
function applyMerges(cells: Map<string, XlsxCell>, merges: XlsxMerge[]) {
  merges.forEach((merge) => {
    const startRef = `${columnIndexToLabel(merge.startColumn)}${
      merge.startRow
    }`;
    const root = cells.get(startRef);
    if (root) {
      root.colSpan = merge.endColumn - merge.startColumn + 1;
      root.rowSpan = merge.endRow - merge.startRow + 1;
    }

    for (let row = merge.startRow; row <= merge.endRow; row += 1) {
      for (
        let column = merge.startColumn;
        column <= merge.endColumn;
        column += 1
      ) {
        if (row === merge.startRow && column === merge.startColumn) continue;
        const ref = `${columnIndexToLabel(column)}${row}`;
        const cell = cells.get(ref);
        if (cell) {
          cell.hiddenByMerge = true;
        } else {
          cells.set(ref, {
            ref,
            rowIndex: row,
            columnIndex: column,
            value: '',
            hiddenByMerge: true,
          });
        }
      }
    }
  });
}

/** 解析完整加载到内存中的 XLSX 工作表及其浮动对象。 */
export function parseMaterializedXlsxSheet(
  xml: string,
  sheetInfo: Pick<XlsxSheet, 'id' | 'name' | 'path'>,
  sharedStrings: string[],
  styleBook: StyleBook,
  packageState: MaterializedXlsxPackageState,
): XlsxSheet {
  // 先读取真实单元格，再补齐空白单元格，确保渲染层能按矩阵方式稳定生成表格。
  const doc = parseXml(xml);
  const sheetNode = doc.documentElement;
  const range = attr(childByLocalName(sheetNode, 'dimension'), 'ref');
  const parsedRange = parseRange(range);
  const metrics = readSheetMetrics(sheetNode, styleBook);
  const cells = new Map<string, XlsxCell>();
  const cellImagePlacements: WpsCellImagePlacement[] = [];
  let maxRow = parsedRange?.endRow ?? 0;
  let maxColumn = parsedRange?.endColumn ?? 0;

  descendantsByLocalName(sheetNode, 'c').forEach((cellNode) => {
    const ref = attr(cellNode, 'r') ?? 'A1';
    const address = parseCellRef(ref);
    const styleId = attr(cellNode, 's')
      ? Number(attr(cellNode, 's'))
      : undefined;
    const value = readCellValue(cellNode, sharedStrings);
    const formula = textContent(childByLocalName(cellNode, 'f'));
    const cellImageId = readWpsCellImageId(formula, value.value);
    const cell: XlsxCell = applyStaticXlsxFormulaHyperlink({
      ref,
      rowIndex: address.row,
      columnIndex: address.column,
      value: cellImageId ? '' : value.value,
      rawValue: value.rawValue,
      type: attr(cellNode, 't'),
      styleId,
      style: resolveStyle(styleId, styleBook),
      formula: formula || undefined,
    });
    if (cellImageId) {
      cellImagePlacements.push({
        imageId: cellImageId,
        row: address.row,
        column: address.column,
      });
    }
    cells.set(ref, cell);
    maxRow = Math.max(maxRow, address.row);
    maxColumn = Math.max(maxColumn, address.column);
  });

  const sheetRelationshipPath = sheetInfo.path
    .replace(/^xl\/worksheets\//, 'xl/worksheets/_rels/')
    .concat('.rels');
  const sheetRelationships =
    packageState.relationships[sheetRelationshipPath] ?? {};
  const hyperlinks = descendantsByLocalName(sheetNode, 'hyperlink').flatMap(
    (node) => {
      const attributes = new Map<string, string>();
      Array.from(node.attributes).forEach((attribute) =>
        attributes.set(attribute.name, attribute.value),
      );
      const hyperlink = parseXlsxHyperlink(attributes, sheetRelationships);
      return hyperlink ? [hyperlink] : [];
    },
  );
  applyXlsxHyperlinkRanges(cells, hyperlinks);
  const annotations = parseMaterializedXlsxComments({
    sheetId: sheetInfo.id,
    relationships: sheetRelationships,
    entries: packageState.entries,
  });
  annotations.forEach((annotation) => {
    let cell = cells.get(annotation.ref);
    if (!cell) {
      cell = {
        ref: annotation.ref,
        rowIndex: annotation.row,
        columnIndex: annotation.column,
        value: '',
      };
      cells.set(annotation.ref, cell);
    }
    cell.annotation ??= annotation;
    maxRow = Math.max(maxRow, annotation.row);
    maxColumn = Math.max(maxColumn, annotation.column);
  });
  const tableSemantics = parseMaterializedXlsxTables({
    sheetNode,
    relationships: sheetRelationships,
    entries: packageState.entries,
  });
  tableSemantics.tables.forEach((table) => {
    maxRow = Math.max(maxRow, table.range.endRow);
    maxColumn = Math.max(maxColumn, table.range.endColumn);
  });
  const conditionalFormatting = parseMaterializedXlsxConditionalFormatting(
    sheetNode,
    styleBook,
    packageState.theme,
  );
  conditionalFormatting.forEach((rule) => {
    rule.ranges.forEach((ruleRange) => {
      maxRow = Math.max(maxRow, ruleRange.endRow);
      maxColumn = Math.max(maxColumn, ruleRange.endColumn);
    });
  });

  const merges = readMerges(sheetNode);
  merges.forEach((merge) => {
    maxRow = Math.max(maxRow, merge.endRow);
    maxColumn = Math.max(maxColumn, merge.endColumn);
  });

  const drawingBounds = readDrawingBounds(
    sheetNode,
    sheetInfo.path,
    packageState,
    metrics,
  );
  if (drawingBounds) {
    // 图片/图表可能锚定在没有单元格内容的区域，需要扩展表格范围保证它们可见。
    maxRow = Math.max(maxRow, drawingBounds.maxRow);
    maxColumn = Math.max(maxColumn, drawingBounds.maxColumn);
  }
  const olePreviewBounds = readOlePreviewBounds(sheetNode);
  if (olePreviewBounds) {
    maxRow = Math.max(maxRow, olePreviewBounds.maxRow);
    maxColumn = Math.max(maxColumn, olePreviewBounds.maxColumn);
  }

  maxRow = Math.min(Math.max(maxRow, 1), MAX_RENDERED_EMPTY_ROWS);
  maxColumn = Math.min(Math.max(maxColumn, 1), MAX_RENDERED_EMPTY_COLUMNS);

  const semanticCells = [...cells.values()];
  applySpreadsheetTableSemantics(
    semanticCells,
    {
      startRow: 1,
      endRow: maxRow,
      startColumn: 1,
      endColumn: maxColumn,
    },
    tableSemantics.tables,
    tableSemantics.autoFilter,
  );
  applyConditionalFormatting(semanticCells, conditionalFormatting);
  semanticCells.forEach((cell) => cells.set(cell.ref, cell));
  applyMerges(cells, merges);

  const rowHeights = new Map<number, number>();
  const customRowHeights = new Set<number>();
  const hiddenRows = new Set<number>();
  descendantsByLocalName(sheetNode, 'row').forEach((rowNode) => {
    const rowIndex = Number(attr(rowNode, 'r') ?? 0);
    if (!rowIndex) return;
    const sourceHeight = attr(rowNode, 'ht');
    if (sourceHeight !== undefined) {
      rowHeights.set(
        rowIndex,
        pointToPx(Number(sourceHeight), metrics.defaultRowHeight),
      );
    }
    const customHeight = attr(rowNode, 'customHeight');
    if (customHeight === '1' || customHeight === 'true') {
      customRowHeights.add(rowIndex);
    }
    const hidden = attr(rowNode, 'hidden');
    if (hidden === '1' || hidden === 'true') hiddenRows.add(rowIndex);
  });
  const columns = readColumns(sheetNode, maxColumn, metrics);

  const rows: XlsxRow[] = Array.from({ length: maxRow }, (_, rowOffset) => {
    const rowIndex = rowOffset + 1;
    return {
      index: rowIndex,
      height: rowHeights.get(rowIndex) ?? metrics.defaultRowHeight,
      customHeight: customRowHeights.has(rowIndex),
      hidden: hiddenRows.has(rowIndex),
      cells: Array.from({ length: maxColumn }, (_, columnOffset) => {
        const columnIndex = columnOffset + 1;
        const ref = `${columnIndexToLabel(columnIndex)}${rowIndex}`;
        return (
          cells.get(ref) ?? {
            ref,
            rowIndex,
            columnIndex,
            value: '',
          }
        );
      }),
    };
  });

  const drawingImages = readSheetImages(
    sheetNode,
    sheetInfo.path,
    packageState,
    columns,
    rowHeights,
    metrics,
  );
  const olePreviewImages = readSheetOlePreviewImages(
    sheetNode,
    sheetInfo.path,
    packageState,
    columns,
    rowHeights,
    metrics,
  );
  return {
    ...sheetInfo,
    defaultColumnWidth: metrics.defaultColumnWidth,
    defaultRowHeight: metrics.defaultRowHeight,
    range,
    rowCount: maxRow,
    columnCount: maxColumn,
    columns,
    rows,
    merges,
    images: [
      ...mergeXlsxPreviewImages(drawingImages, olePreviewImages),
      ...readSheetCellImages(
        cellImagePlacements,
        merges,
        packageState,
        columns,
        rowHeights,
        metrics,
      ),
    ],
    charts: readSheetCharts(
      sheetNode,
      sheetInfo.path,
      packageState,
      columns,
      rowHeights,
      metrics,
    ),
    hyperlinks,
    pane: parseMaterializedXlsxPane(sheetNode),
    tables: tableSemantics.tables.length ? tableSemantics.tables : undefined,
    autoFilter: tableSemantics.autoFilter,
    annotations: annotations.length ? annotations : undefined,
    conditionalFormatting: conditionalFormatting.length
      ? conditionalFormatting
      : undefined,
  };
}
