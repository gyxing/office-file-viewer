import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import { readBinary, readXml } from '../../shared/ooxml/archive';
import { parseOfficeChartXml } from '../../shared/ooxml/charts';
import {
  collectMedia,
  resolvePackageMediaRef,
  type OfficeRelationship,
} from '../../shared/ooxml/media';
import { readRelationships } from '../../shared/ooxml/relationships';
import {
  readOfficeTheme,
  resolveOfficeThemeColor,
  type OfficeTheme,
} from '../../shared/ooxml/theme';
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
import type { SpreadsheetDiagonalBorder } from '../spreadsheet/types';
import {
  parseWpsCellImageDefinitions,
  readWpsCellImageId,
  type WpsCellImagePlacement,
} from '../spreadsheet/wpsCellImages';
import { loadXlsxEntries } from './archive';
import { xlsxImageBytesToDataUrl } from './createXlsxImageResource';
import { mergeXlsxPreviewImages } from './loadXlsxOlePreviewImages';
import type {
  XlsxCell,
  XlsxCellStyle,
  XlsxChart,
  XlsxColumn,
  XlsxImage,
  XlsxMerge,
  XlsxRow,
  XlsxSheet,
  XlsxWorkbook,
} from './types';

/** XLSX 单元格格式引用的字体、填充、边框和对齐。 */
export type ParsedStyle = {
  /** 当前单元格格式引用的字体索引。 */
  fontId?: number;
  /** 当前单元格格式引用的填充索引。 */
  fillId?: number;
  /** 当前单元格格式引用的边框索引。 */
  borderId?: number;
  /** 解析出的单元格水平、垂直对齐与换行配置。 */
  alignment?: XlsxCellStyle;
};

/** XLSX 工作簿共享的字体、填充、边框和单元格格式表。 */
export type StyleBook = {
  /** 按字体索引排列的文字样式。 */
  fonts: XlsxCellStyle[];
  /** 按填充索引排列的背景样式。 */
  fills: Array<Pick<XlsxCellStyle, 'backgroundColor'>>;
  /** 按边框索引排列的单元格边框样式。 */
  borders: Array<
    Pick<
      XlsxCellStyle,
      | 'border'
      | 'borderTop'
      | 'borderRight'
      | 'borderBottom'
      | 'borderLeft'
      | 'borderColor'
      | 'borderWidth'
      | 'diagonalBorder'
    >
  >;
  /** 按样式索引排列的单元格格式引用。 */
  styles: ParsedStyle[];
};

/** 完成解析的零基单元格坐标。 */
type CellAddress = {
  /** 单元格的零基行索引。 */
  row: number;
  /** 单元格的零基列索引。 */
  column: number;
};

/** 枚举 XLSX 解析可能处于的状态。 */
type XlsxPackageState = {
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

/** Excel 默认行高，单位为磅。 */
const DEFAULT_ROW_HEIGHT_POINTS = 15;
/** Excel 默认列宽，单位为标准化渲染像素。 */
const DEFAULT_COLUMN_WIDTH = 64;
/** Excel 默认行高，单位为标准化渲染像素。 */
const DEFAULT_ROW_HEIGHT = 20;
/** 解析工作表尺寸时允许补充渲染的最大空白行数。 */
const MAX_RENDERED_EMPTY_ROWS = 200;
/** 解析工作表尺寸时允许补充渲染的最大空白列数。 */
const MAX_RENDERED_EMPTY_COLUMNS = 80;

/** Excel 主题颜色索引对应的主题槽位顺序。 */
const THEME_COLOR_INDEXES = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
];
/** Excel 内置颜色索引对应的 RGB 颜色表。 */
const INDEXED_COLORS = [
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '800000',
  '008000',
  '000080',
  '808000',
  '800080',
  '008080',
  'C0C0C0',
  '808080',
  '9999FF',
  '993366',
  'FFFFCC',
  'CCFFFF',
  '660066',
  'FF8080',
  '0066CC',
  'CCCCFF',
  '000080',
  'FF00FF',
  'FFFF00',
  '00FFFF',
  '800080',
  '800000',
  '008080',
  '0000FF',
  '00CCFF',
  'CCFFFF',
  'CCFFCC',
  'FFFF99',
  '99CCFF',
  'FF99CC',
  'CC99FF',
  'FFCC99',
  '3366FF',
  '33CCCC',
  '99CC00',
  'FFCC00',
  'FF9900',
  'FF6600',
  '666699',
  '969696',
  '003366',
  '339966',
  '003300',
  '333300',
  '993300',
  '993366',
  '333399',
  '333333',
];

/** 还原 XLSX 行高和列宽所需的默认度量。 */
type SheetMetrics = {
  /** 工作表默认列宽，单位为标准化渲染像素。 */
  defaultColumnWidth: number;
  /** 工作表默认行高，单位为标准化渲染像素。 */
  defaultRowHeight: number;
  /** Normal 字体中数字 0 的最大像素宽度，用于还原 OOXML 字符列宽。 */
  maxDigitWidth: number;
};

// XLSX 中工作表、drawing、chart、media 分散在不同 XML，通过关系表统一解析引用路径。
function buildPackageState(entries: OfficeEntryMap): XlsxPackageState {
  const relationships: XlsxPackageState['relationships'] = {};

  for (const [path, value] of entries) {
    if (typeof value === 'string' && path.endsWith('.rels')) {
      relationships[path] = readRelationships(value, path);
    }
  }

  const media = collectMedia(entries, 'xl/media/');
  // 浏览器不原生支持 metafile；物化路径在建包时一次转换，按需路径则延迟到图片可见时转换。
  for (const [path] of entries) {
    if (!/^xl\/media\/.*\.(?:emf|wmf)$/i.test(path)) continue;
    const binary = readBinary(entries, path);
    if (!binary) continue;
    try {
      const dataUrl = xlsxImageBytesToDataUrl(path, binary);
      media.byPath[path] = dataUrl;
      media.byName[path.split('/').pop() ?? path] = dataUrl;
    } catch {
      // 个别损坏的 metafile 保留原资源，不能阻断其余工作表解析。
    }
  }

  return {
    entries,
    relationships,
    mediaByPath: media.byPath,
    mediaByName: media.byName,
    theme: readOfficeTheme(readXml(entries, 'xl/theme/theme1.xml')),
  };
}

/** 解码 `decodeMojibake` 接收的源数据。 */
export function decodeMojibake(value: string) {
  if (!/[ÃÂäåæçèé]|�|锟|鎬|宸|濮|韬|骞|涓|鍚|煡/.test(value)) {
    return value;
  }

  try {
    const bytes = new Uint8Array(
      Array.from(value, (char) => char.charCodeAt(0) & 0xff),
    );
    const decoded = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
    return decoded.includes('\uFFFD') ? value : decoded;
  } catch {
    return value;
  }
}

function readPlainText(node: Element | null | undefined) {
  if (!node) return '';
  return decodeMojibake(
    descendantsByLocalName(node, 't')
      .map((item) => textContent(item))
      .join(''),
  );
}

function readSharedStrings(xml: string) {
  if (!xml) return [];
  const doc = parseXml(xml);
  return childrenByLocalName(doc.documentElement, 'si').map(readPlainText);
}

/** 将 Excel 列标签转换为零基列索引。 */
export function columnLabelToIndex(label: string) {
  return label
    .split('')
    .reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0);
}

/** 将零基列索引转换为 Excel 列标签。 */
export function columnIndexToLabel(index: number) {
  let value = index;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

/** 解析 A1 单元格引用并返回零基坐标。 */
export function parseCellRef(ref: string): CellAddress {
  const match = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return { row: 1, column: 1 };
  return {
    row: Number(match[2]),
    column: columnLabelToIndex(match[1].toUpperCase()),
  };
}

/** 解析 A1 区域引用并返回行列边界。 */
export function parseRange(range?: string) {
  if (!range) return undefined;
  const [start, end = start] = range.replace(/\$/g, '').split(':');
  const startCell = parseCellRef(start);
  const endCell = parseCellRef(end);
  return {
    startRow: startCell.row,
    startColumn: startCell.column,
    endRow: endCell.row,
    endColumn: endCell.column,
  };
}

/** 将输入标准化为 `normalizeHexColor` 返回的结构。 */
function normalizeHexColor(value?: string) {
  if (!value) return undefined;
  const normalized = value.replace(/^#/, '');
  if (!/^[0-9a-f]{6}$|^[0-9a-f]{8}$/i.test(normalized)) return undefined;
  return `#${normalized.length === 8 ? normalized.slice(2) : normalized}`;
}

function clamp255(value: number) {
  return Math.max(0, Math.min(255, Math.round(value)));
}

function hexToRgb(hex: string) {
  const normalized = hex.replace('#', '');
  const value = Number.parseInt(normalized, 16);
  return {
    r: (value >> 16) & 255,
    g: (value >> 8) & 255,
    b: value & 255,
  };
}

function rgbToHex(r: number, g: number, b: number) {
  return `#${[r, g, b]
    .map((value) => clamp255(value).toString(16).padStart(2, '0'))
    .join('')}`;
}

/** 把 `applyTint` 对应的规则应用到目标对象。 */
function applyTint(hex: string | undefined, tintValue?: string) {
  if (!hex || tintValue === undefined) return hex;
  const tint = Number(tintValue);
  if (!Number.isFinite(tint) || tint === 0) return hex;
  const { r, g, b } = hexToRgb(hex);
  if (tint < 0) {
    const ratio = 1 + tint;
    return rgbToHex(r * ratio, g * ratio, b * ratio);
  }
  return rgbToHex(
    r + (255 - r) * tint,
    g + (255 - g) * tint,
    b + (255 - b) * tint,
  );
}

function parseColor(node: Element | null | undefined, theme: OfficeTheme) {
  if (!node) return undefined;
  if (attr(node, 'auto') === '1') return '#000000';
  const rgb = attr(node, 'rgb');
  const themeIndex = attr(node, 'theme');
  const indexed = attr(node, 'indexed');
  const base =
    normalizeHexColor(rgb) ??
    resolveOfficeThemeColor(
      themeIndex ? THEME_COLOR_INDEXES[Number(themeIndex)] : undefined,
      theme,
    ) ??
    normalizeHexColor(indexed ? INDEXED_COLORS[Number(indexed)] : undefined);
  return applyTint(base, attr(node, 'tint'));
}

function parseBorderStyle(
  node: Element | null | undefined,
  theme: OfficeTheme,
) {
  if (!node) return undefined;
  const style = attr(node, 'style');
  if (!style) return undefined;
  const color = parseColor(childByLocalName(node, 'color'), theme);
  return {
    style,
    color,
    width:
      style === 'hair'
        ? 0.5
        : style === 'medium' || style === 'double'
        ? 2
        : style === 'thick'
        ? 3
        : 1,
  };
}

function borderToCss(border?: ReturnType<typeof parseBorderStyle>) {
  if (!border) return undefined;
  const cssStyle =
    border.style === 'dashed' ||
    border.style === 'dashDot' ||
    border.style === 'dashDotDot' ||
    border.style === 'slantDashDot'
      ? 'dashed'
      : border.style === 'dotted'
      ? 'dotted'
      : border.style === 'double'
      ? 'double'
      : 'solid';
  return `${border.width ?? 1}px ${cssStyle} ${border.color ?? '#000000'}`;
}

/** 将 OOXML 边框线型限制到预览器支持的稳定枚举。 */
function normalizeDiagonalLineStyle(
  value: string,
): SpreadsheetDiagonalBorder['lineStyle'] {
  return value === 'hair' ||
    value === 'thin' ||
    value === 'medium' ||
    value === 'thick' ||
    value === 'double' ||
    value === 'dotted' ||
    value === 'dashed' ||
    value === 'dashDot' ||
    value === 'dashDotDot' ||
    value === 'slantDashDot'
    ? value
    : 'thin';
}

function pointToCssPx(point?: number) {
  if (!point || !Number.isFinite(point)) return undefined;
  return point * (96 / 72);
}

/** 读取 XLSX 样式表并建立字体、填充和边框索引。 */
export function parseStyles(xml: string, theme: OfficeTheme): StyleBook {
  if (!xml) return { fonts: [], fills: [], borders: [], styles: [] };
  const doc = parseXml(xml);
  const styleSheet = doc.documentElement;

  const fontsNode = childByLocalName(styleSheet, 'fonts');
  const fonts = childrenByLocalName(fontsNode, 'font').map(
    (fontNode): XlsxCellStyle => ({
      bold: Boolean(childByLocalName(fontNode, 'b')),
      italic: Boolean(childByLocalName(fontNode, 'i')),
      underline: Boolean(childByLocalName(fontNode, 'u')),
      color: parseColor(childByLocalName(fontNode, 'color'), theme),
      fontSize: pointToCssPx(
        Number(attr(childByLocalName(fontNode, 'sz'), 'val') ?? 0),
      ),
      fontFamily:
        attr(childByLocalName(fontNode, 'name'), 'val') ??
        attr(childByLocalName(fontNode, 'family'), 'val') ??
        attr(childByLocalName(fontNode, 'charset'), 'val') ??
        undefined,
    }),
  );

  const fillsNode = childByLocalName(styleSheet, 'fills');
  const fills = childrenByLocalName(fillsNode, 'fill').map((fillNode) => {
    const pattern = childByLocalName(fillNode, 'patternFill');
    return {
      backgroundColor: parseColor(childByLocalName(pattern, 'fgColor'), theme),
    };
  });

  const bordersNode = childByLocalName(styleSheet, 'borders');
  const borders = childrenByLocalName(bordersNode, 'border').map(
    (borderNode) => {
      const left = parseBorderStyle(
        childByLocalName(borderNode, 'left'),
        theme,
      );
      const right = parseBorderStyle(
        childByLocalName(borderNode, 'right'),
        theme,
      );
      const top = parseBorderStyle(childByLocalName(borderNode, 'top'), theme);
      const bottom = parseBorderStyle(
        childByLocalName(borderNode, 'bottom'),
        theme,
      );
      const diagonal = parseBorderStyle(
        childByLocalName(borderNode, 'diagonal'),
        theme,
      );
      const diagonalUp =
        attr(borderNode, 'diagonalUp') === '1' ||
        attr(borderNode, 'diagonalUp') === 'true';
      const diagonalDown =
        attr(borderNode, 'diagonalDown') === '1' ||
        attr(borderNode, 'diagonalDown') === 'true';
      const color = left?.color ?? right?.color ?? top?.color ?? bottom?.color;
      const width = left?.width ?? right?.width ?? top?.width ?? bottom?.width;
      return {
        border: Boolean(left || right || top || bottom),
        borderTop: borderToCss(top),
        borderRight: borderToCss(right),
        borderBottom: borderToCss(bottom),
        borderLeft: borderToCss(left),
        borderColor: color,
        borderWidth: width,
        diagonalBorder:
          diagonal && (diagonalUp || diagonalDown)
            ? {
                up: diagonalUp,
                down: diagonalDown,
                color: diagonal.color ?? '#000000',
                width: diagonal.width ?? 1,
                lineStyle: normalizeDiagonalLineStyle(diagonal.style),
              }
            : undefined,
      };
    },
  );

  const cellXfs = childByLocalName(styleSheet, 'cellXfs');
  const styles = childrenByLocalName(cellXfs, 'xf').map(
    (xfNode): ParsedStyle => {
      const alignment = childByLocalName(xfNode, 'alignment');
      const vertical = attr(alignment, 'vertical');
      const horizontal = attr(alignment, 'horizontal');
      return {
        fontId: Number(attr(xfNode, 'fontId') ?? 0),
        fillId: Number(attr(xfNode, 'fillId') ?? 0),
        borderId: Number(attr(xfNode, 'borderId') ?? 0),
        alignment: {
          horizontalAlign:
            horizontal === 'center' ||
            horizontal === 'right' ||
            horizontal === 'justify'
              ? horizontal
              : horizontal === 'left'
              ? 'left'
              : undefined,
          verticalAlign:
            vertical === 'top'
              ? 'top'
              : vertical === 'bottom'
              ? 'bottom'
              : vertical === 'center'
              ? 'middle'
              : undefined,
          wrapText: attr(alignment, 'wrapText') === '1',
          shrinkToFit: attr(alignment, 'shrinkToFit') === '1',
        },
      };
    },
  );

  return { fonts, fills, borders, styles };
}

/** 解析并确定 `resolveStyle` 对应的引用或配置。 */
export function resolveStyle(
  styleId: number | undefined,
  styleBook: StyleBook,
): XlsxCellStyle | undefined {
  if (styleId === undefined) return undefined;
  const style = styleBook.styles[styleId];
  if (!style) return undefined;

  const resolved: XlsxCellStyle = {
    ...styleBook.fonts[style.fontId ?? 0],
    ...styleBook.fills[style.fillId ?? 0],
    ...styleBook.borders[style.borderId ?? 0],
    ...style.alignment,
  };

  return Object.fromEntries(
    Object.entries(resolved).filter(
      ([, value]) => value !== undefined && value !== false,
    ),
  ) as XlsxCellStyle;
}

/** 将 Excel 字符列宽换算为标准化渲染像素。 */
export function excelWidthToPx(
  width?: number,
  fallback = DEFAULT_COLUMN_WIDTH,
  maxDigitWidth = 7,
) {
  if (!width || !Number.isFinite(width)) return fallback;
  const safeDigitWidth = Math.max(1, Math.floor(maxDigitWidth));
  return Math.max(
    1,
    Math.floor(
      ((width * 256 + Math.floor(128 / safeDigitWidth)) / 256) * safeDigitWidth,
    ),
  );
}

// Excel 使用 Normal 字体的最大数字宽度换算字符列宽，CJK 字体的数字通常比西文字体估值更宽。
/** 识别中日韩字体族名称的正则表达式。 */
const CJK_FONT_FAMILY_PATTERN =
  /宋体|新宋体|仿宋|黑体|微软雅黑|simsun|nsimsun|fangsong|simhei|microsoft yahei|ms mincho|mingliu|meiryo|malgun gothic/i;

/** 按 Normal 字体估算 Excel 列宽算法使用的最大数字宽度。 */
export function resolveXlsxMaxDigitWidth(font: XlsxCellStyle | undefined) {
  const fontSize = font?.fontSize ?? 11 * (96 / 72);
  const family = font?.fontFamily?.toLowerCase() ?? '';
  const ratio = CJK_FONT_FAMILY_PATTERN.test(family)
    ? 0.55
    : /arial narrow/.test(family)
    ? 0.45
    : /calibri/.test(family)
    ? 0.48
    : /arial|helvetica|liberation sans/.test(family)
    ? 0.525
    : /courier|consolas|monaco|monospace/.test(family)
    ? 0.6
    : 0.5;
  return Math.max(1, Math.floor(fontSize * ratio));
}

/** 将磅值换算为标准化渲染像素。 */
export function pointToPx(point?: number, fallback = DEFAULT_ROW_HEIGHT) {
  if (!point || !Number.isFinite(point)) return fallback;
  return Math.max(1, Math.round(point * (96 / 72)));
}

/** 获取 `getColumnWidth` 返回的数据。 */
function getColumnWidth(
  columns: XlsxColumn[],
  columnIndex: number,
  metrics: SheetMetrics,
) {
  return columns[columnIndex - 1]?.width ?? metrics.defaultColumnWidth;
}

/** 获取 `getRowHeight` 返回的数据。 */
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

/** 解析并确定 `resolveMediaRef` 对应的引用或配置。 */
function resolveMediaRef(
  target: string | undefined,
  packageState: XlsxPackageState,
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
  packageState: XlsxPackageState,
) {
  if (!target) return undefined;
  const normalized = target.replace(/^\.\.\//, '');
  return packageState.entries.get(normalized) ? normalized : target;
}

function readDrawingXml(
  sheetNode: Element,
  sheetPath: string,
  packageState: XlsxPackageState,
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
  packageState: XlsxPackageState,
) {
  const drawing = readDrawingXml(sheetNode, sheetPath, packageState);
  if (!drawing) return undefined;
  const drawingDoc = parseXml(drawing.drawingXml);
  let maxRow = 0;
  let maxColumn = 0;
  childrenByLocalName(drawingDoc.documentElement, 'twoCellAnchor').forEach(
    (anchorNode) => {
      const to = readAnchorPoint(childByLocalName(anchorNode, 'to'));
      maxRow = Math.max(maxRow, to.row);
      maxColumn = Math.max(maxColumn, to.column);
    },
  );
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
  packageState: XlsxPackageState,
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

  return childrenByLocalName(drawingDoc.documentElement, 'twoCellAnchor')
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

      const startPoint = readAnchorPoint(childByLocalName(anchorNode, 'from'));
      const endPoint = readAnchorPoint(childByLocalName(anchorNode, 'to'));
      const start = anchorPosition(startPoint, columns, rowHeights, metrics);
      const end = anchorPosition(endPoint, columns, rowHeights, metrics);
      const chart = parseOfficeChartXml(xml, packageState.theme);
      const name = attr(descendantByLocalName(anchorNode, 'cNvPr'), 'name');

      return {
        id: `${drawing.drawingPath}-chart-${index + 1}`,
        title: name,
        chart,
        from: startPoint,
        to: endPoint,
        x: start.x,
        y: start.y,
        width: Math.max(1, end.x - start.x),
        height: Math.max(1, end.y - start.y),
      };
    })
    .filter(Boolean) as XlsxChart[];
}

function readSheetImages(
  sheetNode: Element,
  sheetPath: string,
  packageState: XlsxPackageState,
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

  return childrenByLocalName(drawingDoc.documentElement, 'twoCellAnchor')
    .map((anchorNode, index): XlsxImage | undefined => {
      const from = readAnchorPoint(childByLocalName(anchorNode, 'from'));
      const to = readAnchorPoint(childByLocalName(anchorNode, 'to'));
      const blip = descendantByLocalName(anchorNode, 'blip');
      const embed = attr(blip, 'r:embed') ?? attr(blip, 'embed');
      const target = embed ? drawingRels[embed]?.target : undefined;
      const src = resolveMediaRef(target, packageState);
      if (!src) return undefined;

      const start = anchorPosition(from, columns, rowHeights, metrics);
      const end = anchorPosition(to, columns, rowHeights, metrics);
      const name = attr(descendantByLocalName(anchorNode, 'cNvPr'), 'name');

      return {
        id: `${drawing.drawingPath}-${index + 1}`,
        name,
        alt: name,
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

/** 读取嵌入对象 objectPr 记录的预览图及单元格锚点。 */
function readSheetOlePreviewImages(
  sheetNode: Element,
  sheetPath: string,
  packageState: XlsxPackageState,
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
  packageState: XlsxPackageState,
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

/** 把 `applyMerges` 对应的规则应用到目标对象。 */
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

function readSheet(
  xml: string,
  sheetInfo: Pick<XlsxSheet, 'id' | 'name' | 'path'>,
  sharedStrings: string[],
  styleBook: StyleBook,
  packageState: XlsxPackageState,
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
    const cell: XlsxCell = {
      ref,
      rowIndex: address.row,
      columnIndex: address.column,
      value: cellImageId ? '' : value.value,
      rawValue: value.rawValue,
      type: attr(cellNode, 't'),
      styleId,
      style: resolveStyle(styleId, styleBook),
      formula: formula || undefined,
    };
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

  const merges = readMerges(sheetNode);
  merges.forEach((merge) => {
    maxRow = Math.max(maxRow, merge.endRow);
    maxColumn = Math.max(maxColumn, merge.endColumn);
  });

  const drawingBounds = readDrawingBounds(
    sheetNode,
    sheetInfo.path,
    packageState,
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
  };
}

function throwIfXlsxParseAborted(signal?: AbortSignal) {
  if (!signal?.aborted) return;
  const error = new Error('XLSX 解析已取消');
  error.name = 'AbortError';
  throw error;
}

/** 在工作表边界让出主线程，使大工作簿切换或卸载可以及时取消。 */
async function xlsxParseCheckpoint(signal?: AbortSignal) {
  throwIfXlsxParseAborted(signal);
  await new Promise<void>((resolve) => {
    setTimeout(resolve, 0);
  });
  throwIfXlsxParseAborted(signal);
}

/** 解析 XLSX 包并返回标准工作簿模型。 */
export async function parseXlsx(
  file: File,
  signal?: AbortSignal,
): Promise<XlsxWorkbook> {
  // sharedStrings 和 styles 是全工作簿共享数据，先解析后再逐个 sheet 套用。
  throwIfXlsxParseAborted(signal);
  const entries = await loadXlsxEntries(file, { signal });
  await xlsxParseCheckpoint(signal);
  const packageState = buildPackageState(entries);
  const workbookXml = readXml(entries, 'xl/workbook.xml');
  const workbookRels =
    packageState.relationships['xl/_rels/workbook.xml.rels'] ?? {};
  const sharedStrings = readSharedStrings(
    readXml(entries, 'xl/sharedStrings.xml'),
  );
  const styleBook = parseStyles(
    readXml(entries, 'xl/styles.xml'),
    packageState.theme,
  );
  const workbookDoc = parseXml(workbookXml);
  const sheetEntries = childrenByLocalName(
    childByLocalName(workbookDoc.documentElement, 'sheets'),
    'sheet',
  )
    .map((node, sourceIndex) => ({ node, sourceIndex }))
    .filter(({ node }) => {
      // Excel/WPS 的隐藏工作表不属于用户可见预览内容。
      const state = attr(node, 'state');
      return state !== 'hidden' && state !== 'veryHidden';
    });
  const sheets: XlsxSheet[] = [];
  for (let index = 0; index < sheetEntries.length; index += 1) {
    await xlsxParseCheckpoint(signal);
    const { node: sheetNode, sourceIndex } = sheetEntries[index];
    const relId = attr(sheetNode, 'r:id') ?? attr(sheetNode, 'id') ?? '';
    const rel = workbookRels[relId];
    const path = rel?.target ?? `xl/worksheets/sheet${sourceIndex + 1}.xml`;
    sheets.push(
      readSheet(
        readXml(entries, path),
        {
          id: attr(sheetNode, 'sheetId') ?? String(sourceIndex + 1),
          name: decodeMojibake(
            attr(sheetNode, 'name') ?? `Sheet ${sourceIndex + 1}`,
          ),
          path,
        },
        sharedStrings,
        styleBook,
        packageState,
      ),
    );
  }

  throwIfXlsxParseAborted(signal);
  return { sheets };
}
