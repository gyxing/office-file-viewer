import type { OfficeArtHyperlinkTarget } from '../../shared/officeart';
import type { PortableResource } from '../parsing/protocol/messages';
import type {
  SpreadsheetCell,
  SpreadsheetCellStyle,
  SpreadsheetColumn,
  SpreadsheetDiagonalBorder,
  SpreadsheetMerge,
  SpreadsheetSheet,
  SpreadsheetWorkbook,
} from '../spreadsheet/types';
import {
  applyStaticXlsxFormulaHyperlink,
  applyXlsxHyperlinkRanges,
  createSpreadsheetExternalHyperlink,
  internalHyperlink,
} from '../xlsx/parseXlsxHyperlinks';
import { BIFF8_RECORD } from './biff8/constants';
import { decodeBiff8Formula } from './biff8/formulas';
import { formatBiff8Value } from './biff8/numberFormats';
import { parseBiff8Charts } from './chart/parseCharts';
import { createPortableImageResource } from './drawing/createPortableImageResource';
import {
  parseBiff8Drawings,
  parseBiff8DrawingShapes,
} from './drawing/parseDrawings';
import type { Biff8DrawingShape } from './drawing/types';
import type {
  Biff8BorderStyle,
  Biff8Cell,
  Biff8CellFormat,
  Biff8Font,
  Biff8SheetDescriptor,
  Biff8Workbook,
  Biff8Worksheet,
} from './types';

/** CSS 像素换算使用的每英寸点数。 */
const CSS_DPI = 96;
/** 一英寸包含的 twip 数量。 */
const TWIPS_PER_INCH = 1440;
/** 一英寸包含的磅数。 */
const POINTS_PER_INCH = 72;
/** Excel 缺少列宽时使用的默认像素宽度。 */
const DEFAULT_COLUMN_PIXELS = 64;
/** Excel 缺少行高时使用的默认像素高度。 */
const DEFAULT_ROW_PIXELS = 20;
/** 一英寸包含的 EMU 数量。 */
const EMUS_PER_INCH = 914400;

/** 将宿主无关的 OfficeArt 目标转换为电子表格内部或外部链接。 */
export function spreadsheetHyperlinkFromOfficeArt(
  source: OfficeArtHyperlinkTarget | undefined,
) {
  if (!source) return undefined;
  if (source.target) {
    const target =
      source.location && !source.target.includes('#')
        ? `${source.target}#${source.location}`
        : source.target;
    return createSpreadsheetExternalHyperlink(target);
  }
  return source.location ? internalHyperlink(source.location) : undefined;
}

function columnLabel(index: number) {
  let value = index;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function cellRef(row: number, column: number) {
  return `${columnLabel(column + 1)}${row + 1}`;
}

function twipsToPixels(value: number) {
  // Excel 在屏幕布局中把 BIFF 行高截断到整像素；保留小数会让下方绘图锚点持续累积偏移。
  return Math.max(1, Math.floor((value / TWIPS_PER_INCH) * CSS_DPI));
}

function pointsToPixels(value: number) {
  return (value / POINTS_PER_INCH) * CSS_DPI;
}

/** 估算 Normal 样式字体的最大数字宽度，BIFF 列宽与绘图锚点都依赖该值。 */
function resolveMaxDigitWidth(font: Biff8Font | undefined) {
  if (!font) return 7;
  const family = font.name.trim().toLowerCase();
  const ratio = /arial narrow/.test(family)
    ? 0.45
    : /(arial|helvetica|liberation sans)/.test(family)
    ? 0.56
    : /(courier|consolas|monaco|monospace)/.test(family)
    ? 0.6
    : 0.5;
  const fontSizePixels = pointsToPixels(font.heightTwips / 20);
  return Math.max(1, Math.floor(fontSizePixels * ratio));
}

/** 按 BIFF 规范将字符列宽换算成像素，最大数字宽度来自工作簿 Normal 字体。 */
function columnWidthToPixels(width: number, maxDigitWidth: number) {
  if (!Number.isFinite(width) || width <= 0) return DEFAULT_COLUMN_PIXELS;
  return Math.max(
    1,
    Math.floor(
      ((width * 256 + Math.floor(128 / maxDigitWidth)) / 256) * maxDigitWidth,
    ),
  );
}

/** Excel 系统颜色索引到 CSS 颜色值的映射。 */
const SYSTEM_COLORS: Record<number, string> = {
  0: '#000000',
  1: '#ffffff',
  2: '#ff0000',
  3: '#00ff00',
  4: '#0000ff',
  5: '#ffff00',
  6: '#ff00ff',
  7: '#00ffff',
};

/** 解析并确定 `resolveColor` 对应的引用或配置。 */
function resolveColor(index: number | undefined, palette: string[]) {
  if (index === undefined || index === 64 || index === 0x7fff) return undefined;
  return SYSTEM_COLORS[index] ?? palette[index - 8];
}

function borderToCss(border: Biff8BorderStyle | undefined, palette: string[]) {
  if (!border?.style) return undefined;
  const width = border.style === 2 ? 2 : border.style === 5 ? 3 : 1;
  const lineStyle =
    border.style === 3 || border.style === 8
      ? 'dashed'
      : border.style === 4 || border.style === 7
      ? 'dotted'
      : border.style === 6
      ? 'double'
      : 'solid';
  return `${width}px ${lineStyle} ${
    resolveColor(border.colorIndex, palette) ?? '#000000'
  }`;
}

/** 将 BIFF8 线型编号标准化为渲染层共享的对角边框语义。 */
function diagonalBorderToStyle(
  border: Biff8CellFormat['diagonalBorder'],
  palette: string[],
): SpreadsheetDiagonalBorder | undefined {
  if (!border?.style || (!border.up && !border.down)) return undefined;
  const lineStyle: SpreadsheetDiagonalBorder['lineStyle'] =
    border.style === 7
      ? 'hair'
      : border.style === 2
      ? 'medium'
      : border.style === 5
      ? 'thick'
      : border.style === 6
      ? 'double'
      : border.style === 3 || border.style === 8
      ? 'dashed'
      : border.style === 4
      ? 'dotted'
      : border.style === 9 || border.style === 10
      ? 'dashDot'
      : border.style === 11 || border.style === 12
      ? 'dashDotDot'
      : border.style === 13
      ? 'slantDashDot'
      : 'thin';
  const width =
    border.style === 7
      ? 0.5
      : border.style === 2 ||
        border.style === 8 ||
        border.style === 10 ||
        border.style === 12
      ? 2
      : border.style === 5
      ? 3
      : border.style === 6
      ? 2
      : 1;
  return {
    up: border.up,
    down: border.down,
    color: resolveColor(border.colorIndex, palette) ?? '#000000',
    width,
    lineStyle,
  };
}

function alignmentFromValue(value: number | undefined) {
  if (value === 1) return 'left' as const;
  if (value === 2 || value === 6) return 'center' as const;
  if (value === 3) return 'right' as const;
  if (value === 5 || value === 7) return 'justify' as const;
  return undefined;
}

function verticalAlignmentFromValue(value: number | undefined) {
  if (value === 0) return 'top' as const;
  if (value === 1) return 'middle' as const;
  if (value === 2) return 'bottom' as const;
  return undefined;
}

/**
 * 将 XF 字体编号映射到实际 FONT 记录。
 *
 * BIFF8 固定保留字体编号 4，工作簿不会为它写入 FONT 记录，因此后续
 * 编号必须前移一项；直接按数组下标读取会让所有编号 5 及以上的字体错位。
 */
function resolveBiff8Font(fontIndex: number, fonts: readonly Biff8Font[]) {
  if (fontIndex === 4) return undefined;
  return fonts[fontIndex > 4 ? fontIndex - 1 : fontIndex];
}

/** 解析并确定 `resolveCellStyle` 对应的引用或配置。 */
function resolveCellStyle(
  xf: Biff8CellFormat | undefined,
  workbook: Biff8Workbook,
): SpreadsheetCellStyle | undefined {
  if (!xf) return undefined;
  const { globals } = workbook;
  const font = resolveBiff8Font(xf.fontIndex, globals.fonts);
  const borderTop = borderToCss(xf.topBorder, globals.palette);
  const borderRight = borderToCss(xf.rightBorder, globals.palette);
  const borderBottom = borderToCss(xf.bottomBorder, globals.palette);
  const borderLeft = borderToCss(xf.leftBorder, globals.palette);
  const style: SpreadsheetCellStyle = {
    bold: font?.bold,
    italic: font?.italic,
    underline: font?.underline,
    color: resolveColor(font?.colorIndex, globals.palette),
    fontFamily: font?.name,
    fontSize: font ? pointsToPixels(font.heightTwips / 20) : undefined,
    backgroundColor: xf.fillPattern
      ? resolveColor(xf.fillForegroundColorIndex, globals.palette)
      : undefined,
    horizontalAlign: alignmentFromValue(xf.horizontalAlign),
    verticalAlign: verticalAlignmentFromValue(xf.verticalAlign),
    wrapText: xf.wrapText,
    shrinkToFit: xf.shrinkToFit,
    border: Boolean(borderTop || borderRight || borderBottom || borderLeft),
    borderTop,
    borderRight,
    borderBottom,
    borderLeft,
    diagonalBorder: diagonalBorderToStyle(xf.diagonalBorder, globals.palette),
  };
  return Object.fromEntries(
    Object.entries(style).filter(([, value]) => value !== undefined),
  ) as SpreadsheetCellStyle;
}

/** 计算 `computeUsedRange` 对应的数值。 */
function computeUsedRange(sheet: Biff8Worksheet) {
  let maxRow = Math.max(0, (sheet.dimensions?.lastRowExclusive ?? 0) - 1);
  let maxColumn = Math.max(0, (sheet.dimensions?.lastColumnExclusive ?? 0) - 1);
  for (const cell of sheet.cells) {
    maxRow = Math.max(maxRow, cell.row);
    maxColumn = Math.max(maxColumn, cell.column);
  }
  // ROW/COLINFO 可能只为整行整列保存尺寸或默认样式，不能据此扩大实际内容范围。
  for (const merge of sheet.merges) {
    maxRow = Math.max(maxRow, merge.endRow);
    maxColumn = Math.max(maxColumn, merge.endColumn);
  }
  for (const hyperlink of sheet.hyperlinks) {
    maxRow = Math.max(maxRow, hyperlink.endRow - 1);
    maxColumn = Math.max(maxColumn, hyperlink.endColumn - 1);
  }
  return { maxRow, maxColumn };
}

/** 查找 `findColumnInfo` 对应的目标数据。 */
function findColumnInfo(sheet: Biff8Worksheet, column: number) {
  return sheet.columns.find(
    (item) => column >= item.firstColumn && column <= item.lastColumn,
  );
}

/** 把源数据适配为 `adaptCell` 返回的标准模型。 */
export function adaptBiff8Cell(
  source: Biff8Cell | undefined,
  row: number,
  column: number,
  fallbackXfIndex: number | undefined,
  workbook: Biff8Workbook,
): SpreadsheetCell {
  const xfIndex = source?.xfIndex ?? fallbackXfIndex;
  const xf =
    xfIndex === undefined ? undefined : workbook.globals.cellFormats[xfIndex];
  const rawValue = source?.value;
  const format = xf ? workbook.globals.formats.get(xf.formatIndex) : undefined;
  return applyStaticXlsxFormulaHyperlink({
    ref: cellRef(row, column),
    rowIndex: row + 1,
    columnIndex: column + 1,
    value: source
      ? formatBiff8Value(rawValue ?? null, format, workbook.globals.date1904)
      : '',
    rawValue: source && rawValue !== null ? String(rawValue) : undefined,
    type: source?.cachedType,
    styleId: xfIndex,
    style: resolveCellStyle(xf, workbook),
    formula: source?.formula,
    formulaTokens: source?.formulaTokens,
  });
}

/** 把源数据适配为 `adaptMerges` 返回的标准模型。 */
function adaptMerges(
  sheet: Biff8Worksheet,
  cells: Map<string, SpreadsheetCell>,
) {
  const merges: SpreadsheetMerge[] = [];
  for (const source of sheet.merges) {
    const anchorRef = cellRef(source.startRow, source.startColumn);
    const anchor = cells.get(anchorRef)!;
    anchor.rowSpan = source.endRow - source.startRow + 1;
    anchor.colSpan = source.endColumn - source.startColumn + 1;
    for (let row = source.startRow; row <= source.endRow; row += 1) {
      for (
        let column = source.startColumn;
        column <= source.endColumn;
        column += 1
      ) {
        if (row !== source.startRow || column !== source.startColumn) {
          cells.get(cellRef(row, column))!.hiddenByMerge = true;
        }
      }
    }
    merges.push({
      ref: `${anchorRef}:${cellRef(source.endRow, source.endColumn)}`,
      startRow: source.startRow + 1,
      startColumn: source.startColumn + 1,
      endRow: source.endRow + 1,
      endColumn: source.endColumn + 1,
    });
  }
  return merges;
}

/** 把源数据适配为 `adaptWorksheet` 返回的标准模型。 */
function adaptWorksheet(
  sheet: Biff8Worksheet,
  workbook: Biff8Workbook,
): SpreadsheetSheet {
  const { maxRow, maxColumn } = computeUsedRange(sheet);
  const maxDigitWidth = resolveMaxDigitWidth(workbook.globals.fonts[0]);
  const sourceCells = new Map(
    sheet.cells.map((cell) => [`${cell.row}:${cell.column}`, cell]),
  );
  const cells = new Map<string, SpreadsheetCell>();
  const columns: SpreadsheetColumn[] = Array.from(
    { length: maxColumn + 1 },
    (_, column) => {
      const info = findColumnInfo(sheet, column);
      return {
        index: column + 1,
        label: columnLabel(column + 1),
        width: columnWidthToPixels(
          info?.widthCharacters ?? sheet.defaultColumnWidth,
          maxDigitWidth,
        ),
        hidden: info?.hidden,
      };
    },
  );

  const rowInfoByIndex = new Map(
    sheet.rows.map((rowInfo) => [rowInfo.index, rowInfo]),
  );
  const rows = Array.from({ length: maxRow + 1 }, (_, row) => {
    const rowInfo = rowInfoByIndex.get(row);
    const rowCells = Array.from({ length: maxColumn + 1 }, (_, column) => {
      const columnInfo = findColumnInfo(sheet, column);
      const cell = adaptBiff8Cell(
        sourceCells.get(`${row}:${column}`),
        row,
        column,
        columnInfo?.xfIndex,
        workbook,
      );
      cells.set(cell.ref, cell);
      return cell;
    });
    return {
      index: row + 1,
      height: rowInfo?.heightTwips
        ? twipsToPixels(rowInfo.heightTwips)
        : sheet.defaultRowHeightTwips
        ? twipsToPixels(sheet.defaultRowHeightTwips)
        : DEFAULT_ROW_PIXELS,
      customHeight: rowInfo?.customHeight,
      hidden: rowInfo?.hidden,
      cells: rowCells,
    };
  });
  const merges = adaptMerges(sheet, cells);
  applyXlsxHyperlinkRanges(cells, sheet.hyperlinks);
  const endRef = cellRef(maxRow, maxColumn);
  return {
    id: sheet.descriptor.id,
    name: sheet.descriptor.name,
    path: `/Workbook/${sheet.descriptor.name}`,
    kind: 'worksheet',
    defaultColumnWidth: columnWidthToPixels(
      sheet.defaultColumnWidth,
      maxDigitWidth,
    ),
    defaultRowHeight: sheet.defaultRowHeightTwips
      ? twipsToPixels(sheet.defaultRowHeightTwips)
      : DEFAULT_ROW_PIXELS,
    range: endRef === 'A1' ? 'A1' : `A1:${endRef}`,
    rowCount: maxRow + 1,
    columnCount: maxColumn + 1,
    columns,
    rows,
    merges,
    hyperlinks: sheet.hyperlinks,
    images: [],
    charts: [],
  };
}

function createChartSheetPlaceholder(
  descriptor: Biff8SheetDescriptor,
): SpreadsheetSheet {
  return {
    id: descriptor.id,
    name: descriptor.name,
    path: `/Workbook/${descriptor.name}`,
    kind: 'chart',
    defaultColumnWidth: DEFAULT_COLUMN_PIXELS,
    defaultRowHeight: DEFAULT_ROW_PIXELS,
    range: 'A1',
    rowCount: 1,
    columnCount: 1,
    columns: [{ index: 1, label: 'A', width: DEFAULT_COLUMN_PIXELS }],
    rows: [
      {
        index: 1,
        height: DEFAULT_ROW_PIXELS,
        cells: [{ ref: 'A1', rowIndex: 1, columnIndex: 1, value: '' }],
      },
    ],
    merges: [],
    images: [],
    charts: [],
  };
}

/** 将一个 BIFF8 工作表描述符适配为可独立传输的预览模型。 */
export function adaptBiff8Sheet(
  source: Biff8Workbook,
  descriptor: Biff8SheetDescriptor,
): SpreadsheetSheet | undefined {
  const worksheet = source.worksheets.find(
    (sheet) => sheet.descriptor.id === descriptor.id,
  );
  if (worksheet) return adaptWorksheet(worksheet, source);
  if (descriptor.type !== 'chart') return undefined;
  return createChartSheetPlaceholder(descriptor);
}

/** 把大型 BIFF8 Worksheet 转为稀疏标准模型，不创建完整空矩阵。 */
export function adaptBiff8WorksheetSparse(
  workbook: Biff8Workbook,
  sheet: Biff8Worksheet,
) {
  const { maxRow, maxColumn } = computeUsedRange(sheet);
  const maxDigitWidth = resolveMaxDigitWidth(workbook.globals.fonts[0]);
  const rowCount = maxRow + 1;
  const columnCount = maxColumn + 1;
  const defaultColumnWidth = columnWidthToPixels(
    sheet.defaultColumnWidth,
    maxDigitWidth,
  );
  const defaultRowHeight = sheet.defaultRowHeightTwips
    ? twipsToPixels(sheet.defaultRowHeightTwips)
    : DEFAULT_ROW_PIXELS;
  const columnMetrics = new Map<
    number,
    { index: number; width: number; hidden: boolean }
  >();
  sheet.columns.forEach((column) => {
    for (
      let index = column.firstColumn;
      index <= Math.min(column.lastColumn, maxColumn);
      index += 1
    ) {
      columnMetrics.set(index + 1, {
        index: index + 1,
        width: columnWidthToPixels(column.widthCharacters, maxDigitWidth),
        hidden: Boolean(column.hidden),
      });
    }
  });
  const cells = sheet.cells.map((cell) =>
    adaptBiff8Cell(cell, cell.row, cell.column, undefined, workbook),
  );
  const cellByRef = new Map(cells.map((cell) => [cell.ref, cell]));
  const merges = sheet.merges.map((merge) => {
    const ref = cellRef(merge.startRow, merge.startColumn);
    let root = cellByRef.get(ref);
    if (!root) {
      const columnInfo = findColumnInfo(sheet, merge.startColumn);
      root = adaptBiff8Cell(
        undefined,
        merge.startRow,
        merge.startColumn,
        columnInfo?.xfIndex,
        workbook,
      );
      cells.push(root);
      cellByRef.set(ref, root);
    }
    root.rowSpan = merge.endRow - merge.startRow + 1;
    root.colSpan = merge.endColumn - merge.startColumn + 1;
    return {
      ref: `${ref}:${cellRef(merge.endRow, merge.endColumn)}`,
      startRow: merge.startRow + 1,
      startColumn: merge.startColumn + 1,
      endRow: merge.endRow + 1,
      endColumn: merge.endColumn + 1,
    };
  });
  applyXlsxHyperlinkRanges(cellByRef, sheet.hyperlinks);
  return {
    rowCount,
    columnCount,
    defaultColumnWidth,
    defaultRowHeight,
    rows: sheet.rows.map((row) => ({
      index: row.index + 1,
      height: row.heightTwips
        ? twipsToPixels(row.heightTwips)
        : defaultRowHeight,
      customHeight: row.customHeight,
      hidden: Boolean(row.hidden),
    })),
    columns: [...columnMetrics.values()],
    cells,
    merges,
    hyperlinks: sheet.hyperlinks,
  };
}

/** 把 BIFF8 静态定义名称解码为工作簿内部导航可消费的地址。 */
export function adaptBiff8DefinedNames(globals: Biff8Workbook['globals']) {
  return Object.fromEntries(
    globals.definedNames.flatMap((definedName) => {
      const decoded = decodeBiff8Formula(definedName.tokens, {
        row: 0,
        column: 0,
        definedNames: globals.definedNames,
        sheets: globals.sheets,
        externalSheets: globals.externalSheets,
      });
      const location = decoded.formula?.replace(/^=/, '');
      return location ? [[definedName.name, location] as const] : [];
    }),
  );
}

/** 将 BIFF8 中间模型适配为 XLSX 预览器复用的通用工作簿。 */
export function adaptBiff8Workbook(source: Biff8Workbook): SpreadsheetWorkbook {
  return {
    sheets: source.globals.sheets.flatMap((descriptor) => {
      const sheet = adaptBiff8Sheet(source, descriptor);
      return sheet ? [sheet] : [];
    }),
    definedNames: adaptBiff8DefinedNames(source.globals),
    warnings: source.warnings.length ? source.warnings : undefined,
  };
}

function concatenateChunks(chunks: Uint8Array[]) {
  const length = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const result = new Uint8Array(length);
  let offset = 0;
  chunks.forEach((chunk) => {
    result.set(chunk, offset);
    offset += chunk.length;
  });
  return result;
}

function pointGeometry(
  sheet: SpreadsheetSheet,
  point: {
    /** 锚点单元格的零基行索引。 */
    row: number;
    /** 锚点单元格的零基列索引。 */
    column: number;
    /** 锚点在单元格内的纵向位置比例。 */
    rowFraction: number;
    /** 锚点在单元格内的横向位置比例。 */
    columnFraction: number;
  },
) {
  const x = sheet.columns
    .slice(0, point.column)
    .reduce((sum, column) => sum + (column.hidden ? 0 : column.width), 0);
  const y = sheet.rows
    .slice(0, point.row)
    .reduce((sum, row) => sum + (row.hidden ? 0 : row.height), 0);
  const column = sheet.columns[point.column];
  const row = sheet.rows[point.row];
  const columnWidth = column?.hidden
    ? 0
    : column?.width ?? sheet.defaultColumnWidth ?? DEFAULT_COLUMN_PIXELS;
  const rowHeight = row?.hidden
    ? 0
    : row?.height ?? sheet.defaultRowHeight ?? DEFAULT_ROW_PIXELS;
  return {
    x: x + columnWidth * point.columnFraction,
    y: y + rowHeight * point.rowFraction,
    columnWidth,
    rowHeight,
  };
}

function ensureSheetBounds(
  sheet: SpreadsheetSheet,
  requiredRows: number,
  requiredColumns: number,
) {
  const defaultColumnWidth = sheet.defaultColumnWidth ?? DEFAULT_COLUMN_PIXELS;
  const defaultRowHeight = sheet.defaultRowHeight ?? DEFAULT_ROW_PIXELS;
  while (sheet.columns.length < requiredColumns) {
    const index = sheet.columns.length + 1;
    sheet.columns.push({
      index,
      label: columnLabel(index),
      width: defaultColumnWidth,
    });
    sheet.rows.forEach((row) => {
      row.cells.push({
        ref: `${columnLabel(index)}${row.index}`,
        rowIndex: row.index,
        columnIndex: index,
        value: '',
      });
    });
  }
  while (sheet.rows.length < requiredRows) {
    const index = sheet.rows.length + 1;
    sheet.rows.push({
      index,
      height: defaultRowHeight,
      cells: sheet.columns.map((column) => ({
        ref: `${column.label}${index}`,
        rowIndex: index,
        columnIndex: column.index,
        value: '',
      })),
    });
  }
  sheet.rowCount = sheet.rows.length;
  sheet.columnCount = sheet.columns.length;
  const endRef = `${columnLabel(sheet.columnCount)}${sheet.rowCount}`;
  sheet.range = endRef === 'A1' ? 'A1' : `A1:${endRef}`;
}

/** 接收 XLS 图片资源并返回稳定引用的接口。 */
export type XlsResourceCollector = {
  /** 向当前聚合器添加一个解析结果。 */
  add(resource: PortableResource): Promise<string>;
};

/** 将 OfficeArt 的低三字节 BGR 色值转换为 CSS RGB 颜色。 */
function officeArtColorToCss(value: number | undefined) {
  if (value === undefined) return undefined;
  const red = value & 0xff;
  const green = (value >>> 8) & 0xff;
  const blue = (value >>> 16) & 0xff;
  return `rgb(${red}, ${green}, ${blue})`;
}

/** 转义 SVG 文本节点，避免对象名称破坏生成的图形资源。 */
function escapeSvgText(value: string) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/** 将常见 BIFF8 AutoShape 转成无损缩放的 SVG 图片资源。 */
function createBiff8ShapeResource(
  shape: Biff8DrawingShape,
  width: number,
  height: number,
): PortableResource | undefined {
  if (shape.shapeType !== 1 && shape.shapeType !== 2 && shape.shapeType !== 3) {
    return undefined;
  }
  const fill = officeArtColorToCss(shape.fillColor) ?? 'none';
  const stroke = officeArtColorToCss(shape.lineColor) ?? 'none';
  const strokeWidth =
    shape.lineWidth === undefined
      ? 0
      : Math.max(0, (shape.lineWidth / EMUS_PER_INCH) * CSS_DPI);
  const inset = stroke === 'none' ? 0 : strokeWidth / 2;
  const innerWidth = Math.max(0, width - inset * 2);
  const innerHeight = Math.max(0, height - inset * 2);
  const geometry =
    shape.shapeType === 3
      ? `<ellipse cx="${width / 2}" cy="${height / 2}" rx="${
          innerWidth / 2
        }" ry="${innerHeight / 2}"/>`
      : `<rect x="${inset}" y="${inset}" width="${innerWidth}" height="${innerHeight}"${
          shape.shapeType === 2 ? ` rx="${Math.min(width, height) * 0.12}"` : ''
        }/>`;
  const title = shape.name ? `<title>${escapeSvgText(shape.name)}</title>` : '';
  return {
    id: `xls-shape-resource:${shape.id}`,
    encoding: 'text',
    mimeType: 'image/svg+xml',
    text: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}">${title}<g fill="${fill}" stroke="${stroke}" stroke-width="${strokeWidth}">${geometry}</g></svg>`,
  };
}

/** 解析并附加 XLS 绘图图片，资源的实体化方式由运行环境注入。 */
export async function attachBiff8DrawingImages(
  target: SpreadsheetWorkbook,
  source: Biff8Workbook,
  resources: XlsResourceCollector,
) {
  const groupBytes = concatenateChunks(
    source.globals.drawingGroupRecords.flatMap((record) => record.chunks),
  );
  if (!groupBytes.length) return;
  const warnings = target.warnings ?? [];
  target.warnings = warnings;

  for (const sourceSheet of source.worksheets) {
    const targetSheet = target.sheets.find(
      (sheet) => sheet.id === sourceSheet.descriptor.id,
    );
    if (!targetSheet) continue;
    const sheetBytes = concatenateChunks(
      sourceSheet.drawingRecords
        .filter((record) => record.recordId === BIFF8_RECORD.MSODRAWING)
        .flatMap((record) => record.chunks),
    );
    if (!sheetBytes.length) continue;
    let images: ReturnType<typeof parseBiff8Drawings>;
    let shapes: ReturnType<typeof parseBiff8DrawingShapes>;
    try {
      images = parseBiff8Drawings(groupBytes, sheetBytes, warnings);
      shapes = parseBiff8DrawingShapes(sheetBytes, warnings);
    } catch (error) {
      warnings.push({
        code: 'INVALID_SHEET_DRAWING',
        message: `工作表“${sourceSheet.descriptor.name}”的绘图结构无效：${
          error instanceof Error ? error.message : '未知错误'
        }`,
        sheetName: sourceSheet.descriptor.name,
      });
      continue;
    }
    for (let imageIndex = 0; imageIndex < images.length; imageIndex += 1) {
      const image = images[imageIndex];
      ensureSheetBounds(
        targetSheet,
        image.anchor.to.row + 1,
        image.anchor.to.column + 1,
      );
      try {
        const resource = await createPortableImageResource(
          image,
          `xls:${sourceSheet.descriptor.id}:${image.id}:${imageIndex}`,
        );
        const src = await resources.add(resource.resource);
        warnings.push(
          ...resource.warnings.map((warning) => ({
            ...warning,
            sheetName: warning.sheetName ?? sourceSheet.descriptor.name,
          })),
        );
        const from = pointGeometry(targetSheet, image.anchor.from);
        const to = pointGeometry(targetSheet, image.anchor.to);
        targetSheet.images.push({
          id: image.id,
          name: image.name,
          alt: image.alt,
          src,
          from: {
            row: image.anchor.from.row + 1,
            column: image.anchor.from.column + 1,
            rowOffset: from.rowHeight * image.anchor.from.rowFraction,
            columnOffset: from.columnWidth * image.anchor.from.columnFraction,
          },
          to: {
            row: image.anchor.to.row + 1,
            column: image.anchor.to.column + 1,
            rowOffset: to.rowHeight * image.anchor.to.rowFraction,
            columnOffset: to.columnWidth * image.anchor.to.columnFraction,
          },
          x: from.x,
          y: from.y,
          width: Math.max(1, to.x - from.x),
          height: Math.max(1, to.y - from.y),
          hyperlink: spreadsheetHyperlinkFromOfficeArt(image.hyperlink),
        });
      } catch (error) {
        warnings.push({
          code: 'IMAGE_RENDER_FAILED',
          message: `图片“${image.name ?? image.id}”转换失败：${
            error instanceof Error ? error.message : '未知错误'
          }`,
          sheetName: sourceSheet.descriptor.name,
        });
      }
    }
    // BIFF8 普通形状没有 BLIP 图片数据，统一转成 SVG 后复用现有浮动图片坐标与资源生命周期。
    for (const shape of shapes) {
      if (shape.blipIndex !== undefined) continue;
      if (
        shape.shapeType !== 1 &&
        shape.shapeType !== 2 &&
        shape.shapeType !== 3
      ) {
        continue;
      }
      ensureSheetBounds(
        targetSheet,
        shape.anchor.to.row + 1,
        shape.anchor.to.column + 1,
      );
      const from = pointGeometry(targetSheet, shape.anchor.from);
      const to = pointGeometry(targetSheet, shape.anchor.to);
      const width = Math.max(1, to.x - from.x);
      const height = Math.max(1, to.y - from.y);
      const resource = createBiff8ShapeResource(shape, width, height);
      if (!resource) continue;
      const src = await resources.add(resource);
      targetSheet.images.push({
        id: shape.id,
        name: shape.name,
        alt: shape.name,
        src,
        from: {
          row: shape.anchor.from.row + 1,
          column: shape.anchor.from.column + 1,
          rowOffset: from.rowHeight * shape.anchor.from.rowFraction,
          columnOffset: from.columnWidth * shape.anchor.from.columnFraction,
        },
        to: {
          row: shape.anchor.to.row + 1,
          column: shape.anchor.to.column + 1,
          rowOffset: to.rowHeight * shape.anchor.to.rowFraction,
          columnOffset: to.columnWidth * shape.anchor.to.columnFraction,
        },
        x: from.x,
        y: from.y,
        width,
        height,
        hyperlink: spreadsheetHyperlinkFromOfficeArt(shape.hyperlink),
      });
    }
  }
}

/** 解析并附加内嵌图表与独立图表工作表。 */
export function attachBiff8Charts(
  target: SpreadsheetWorkbook,
  source: Biff8Workbook,
) {
  const warnings = target.warnings ?? [];
  target.warnings = warnings;
  for (const sourceSheet of source.worksheets) {
    if (!sourceSheet.chartSubstreams.length) continue;
    const targetSheet = target.sheets.find(
      (sheet) => sheet.id === sourceSheet.descriptor.id,
    );
    if (!targetSheet) continue;
    const sheetBytes = concatenateChunks(
      sourceSheet.drawingRecords
        .filter((record) => record.recordId === BIFF8_RECORD.MSODRAWING)
        .flatMap((record) => record.chunks),
    );
    const shapes = parseBiff8DrawingShapes(sheetBytes, warnings);
    const charts = parseBiff8Charts(
      source,
      sourceSheet.descriptor,
      sourceSheet.chartSubstreams,
      shapes,
      targetSheet.images,
      sourceSheet,
    );
    charts.forEach((item) => {
      ensureSheetBounds(
        targetSheet,
        item.anchor.to.row + 1,
        item.anchor.to.column + 1,
      );
      const from = pointGeometry(targetSheet, item.anchor.from);
      const to = pointGeometry(targetSheet, item.anchor.to);
      targetSheet.charts.push({
        id: item.id,
        title: item.title,
        chart: item.chart,
        from: {
          row: item.anchor.from.row + 1,
          column: item.anchor.from.column + 1,
          rowOffset: from.rowHeight * item.anchor.from.rowFraction,
          columnOffset: from.columnWidth * item.anchor.from.columnFraction,
        },
        to: {
          row: item.anchor.to.row + 1,
          column: item.anchor.to.column + 1,
          rowOffset: to.rowHeight * item.anchor.to.rowFraction,
          columnOffset: to.columnWidth * item.anchor.to.columnFraction,
        },
        x: from.x,
        y: from.y,
        width: Math.max(1, to.x - from.x),
        height: Math.max(1, to.y - from.y),
        hyperlink: spreadsheetHyperlinkFromOfficeArt(item.hyperlink),
      });
      if (item.previewImageId) {
        targetSheet.images = targetSheet.images.filter(
          (image) => image.id !== item.previewImageId,
        );
      }
      warnings.push(...item.warnings);
    });
  }

  source.chartSheets.forEach((chartSheet) => {
    const targetSheet = target.sheets.find(
      (sheet) => sheet.id === chartSheet.descriptor.id,
    );
    if (!targetSheet) return;
    const chart = parseBiff8Charts(
      source,
      chartSheet.descriptor,
      [chartSheet.substream],
      [],
      [],
    )[0];
    if (!chart) return;
    targetSheet.charts.push({
      id: chart.id,
      title: chart.title,
      chart: chart.chart,
      from: { row: 1, column: 1, rowOffset: 0, columnOffset: 0 },
      to: { row: 1, column: 1, rowOffset: 0, columnOffset: 0 },
      x: 0,
      y: 0,
      width: 960,
      height: 600,
    });
    warnings.push(...chart.warnings);
  });
}
