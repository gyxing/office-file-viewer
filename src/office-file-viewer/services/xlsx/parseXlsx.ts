import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import { readXml } from '../../shared/ooxml/archive';
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
import { loadXlsxEntries } from './archive';
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

/** 描述 XLSX 解析使用的样式参数。 */
type ParsedStyle = {
  /** ParsedStyle 在源文件记录中的数字标识。 */
  fontId?: number;
  /** ParsedStyle 在源文件记录中的数字标识。 */
  fillId?: number;
  /** ParsedStyle 在源文件记录中的数字标识。 */
  borderId?: number;
  /** ParsedStyle 解析出的单元格水平、垂直对齐与换行配置；未提供时使用来源格式或渲染器的默认行为。 */
  alignment?: XlsxCellStyle;
};

/** 描述 StyleBook 在 XLSX 解析中的数据结构。 */
type StyleBook = {
  /** StyleBook 包含的 fonts 有序集合。 */
  fonts: XlsxCellStyle[];
  /** StyleBook 包含的 fills 有序集合。 */
  fills: Array<Pick<XlsxCellStyle, 'backgroundColor'>>;
  /** StyleBook 包含的 borders 有序集合。 */
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
    >
  >;
  /** StyleBook 包含的 styles 有序集合。 */
  styles: ParsedStyle[];
};

/** 描述 CellAddress 在 XLSX 解析中的数据结构。 */
type CellAddress = {
  /** CellAddress 使用的零基行列索引。 */
  row: number;
  /** CellAddress 使用的零基行列索引。 */
  column: number;
};

/** 枚举XLSX 解析可能处于的状态。 */
type XlsxPackageState = {
  /** XlsxPackageState 按压缩包路径索引的文件条目映射。 */
  entries: OfficeEntryMap;
  /** XlsxPackageState 解析得到的 OOXML 关系映射。 */
  relationships: Record<string, Record<string, OfficeRelationship>>;
  /** 按 OOXML 包内路径索引的媒体资源映射。 */
  mediaByPath: Record<string, string>;
  /** 按媒体文件名索引的资源映射。 */
  mediaByName: Record<string, string>;
  /** XlsxPackageState 使用的主题颜色和字体配置。 */
  theme: OfficeTheme;
};

const DEFAULT_COLUMN_WIDTH_CHARACTERS = 8.43;
const DEFAULT_ROW_HEIGHT_POINTS = 15;
const DEFAULT_COLUMN_WIDTH = 64;
const DEFAULT_ROW_HEIGHT = 20;
const MAX_RENDERED_EMPTY_ROWS = 200;
const MAX_RENDERED_EMPTY_COLUMNS = 80;

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

/** 描述 SheetMetrics 在 XLSX 解析中的数据结构。 */
type SheetMetrics = {
  /** SheetMetrics 的 defaultColumnWidth 渲染尺寸，单位为标准化像素。 */
  defaultColumnWidth: number;
  /** SheetMetrics 的 defaultRowHeight 渲染尺寸，单位为标准化像素。 */
  defaultRowHeight: number;
};

// XLSX 中工作表、drawing、chart、media 分散在不同 XML，通过关系表统一解析引用路径。
/** 根据输入构建 `buildPackageState` 返回的标准化结果。 */
function buildPackageState(entries: OfficeEntryMap): XlsxPackageState {
  const relationships: XlsxPackageState['relationships'] = {};

  for (const [path, value] of entries) {
    if (typeof value === 'string' && path.endsWith('.rels')) {
      relationships[path] = readRelationships(value, path);
    }
  }

  const media = collectMedia(entries, 'xl/media/');

  return {
    entries,
    relationships,
    mediaByPath: media.byPath,
    mediaByName: media.byName,
    theme: readOfficeTheme(readXml(entries, 'xl/theme/theme1.xml')),
  };
}

/** 解码 `decodeMojibake` 接收的源数据。 */
function decodeMojibake(value: string) {
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

/** 读取 `readPlainText` 所需的源数据，供 XLSX 解析使用。 */
function readPlainText(node: Element | null | undefined) {
  if (!node) return '';
  return decodeMojibake(
    descendantsByLocalName(node, 't')
      .map((item) => textContent(item))
      .join(''),
  );
}

/** 读取 `readSharedStrings` 所需的源数据，供 XLSX 解析使用。 */
function readSharedStrings(xml: string) {
  if (!xml) return [];
  const doc = parseXml(xml);
  return childrenByLocalName(doc.documentElement, 'si').map(readPlainText);
}

/** 执行 `columnLabelToIndex` 封装的 XLSX 解析处理步骤。 */
function columnLabelToIndex(label: string) {
  return label
    .split('')
    .reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0);
}

/** 执行 `columnIndexToLabel` 封装的 XLSX 解析处理步骤。 */
function columnIndexToLabel(index: number) {
  let value = index;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

/** 解析 `parseCellRef` 接收的数据，并返回 XLSX 解析结果。 */
function parseCellRef(ref: string): CellAddress {
  const match = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return { row: 1, column: 1 };
  return {
    row: Number(match[2]),
    column: columnLabelToIndex(match[1].toUpperCase()),
  };
}

/** 解析 `parseRange` 接收的数据，并返回 XLSX 解析结果。 */
function parseRange(range?: string) {
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

/** 执行 `clamp255` 封装的 XLSX 解析处理步骤。 */
function clamp255(value: number) {
  return Math.max(0, Math.min(255, Math.round(value)));
}

/** 执行 `hexToRgb` 封装的 XLSX 解析处理步骤。 */
function hexToRgb(hex: string) {
  const normalized = hex.replace('#', '');
  const value = Number.parseInt(normalized, 16);
  return {
    r: (value >> 16) & 255,
    g: (value >> 8) & 255,
    b: value & 255,
  };
}

/** 执行 `rgbToHex` 封装的 XLSX 解析处理步骤。 */
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

/** 解析 `parseColor` 接收的数据，并返回 XLSX 解析结果。 */
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

/** 解析 `parseBorderStyle` 接收的数据，并返回 XLSX 解析结果。 */
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

/** 执行 `borderToCss` 封装的 XLSX 解析处理步骤。 */
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

/** 执行 `pointToCssPx` 封装的 XLSX 解析处理步骤。 */
function pointToCssPx(point?: number) {
  if (!point || !Number.isFinite(point)) return undefined;
  return point * (96 / 72);
}

/** 解析 `parseStyles` 接收的数据，并返回 XLSX 解析结果。 */
function parseStyles(xml: string, theme: OfficeTheme): StyleBook {
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
        },
      };
    },
  );

  return { fonts, fills, borders, styles };
}

/** 解析并确定 `resolveStyle` 对应的引用或配置。 */
function resolveStyle(
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

/** 执行 `excelWidthToPx` 封装的 XLSX 解析处理步骤。 */
function excelWidthToPx(width?: number, fallback = DEFAULT_COLUMN_WIDTH) {
  if (!width || !Number.isFinite(width)) return fallback;
  return Math.max(40, Math.round(width * 7 + 5));
}

/** 执行 `pointToPx` 封装的 XLSX 解析处理步骤。 */
function pointToPx(point?: number, fallback = DEFAULT_ROW_HEIGHT) {
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

/** 执行 `anchorPosition` 封装的 XLSX 解析处理步骤。 */
function anchorPosition(
  anchor: {
    /** 当前局部结构 使用的零基行列索引。 */
    row: number;
    /** 当前局部结构 使用的零基行列索引。 */
    column: number;
    /** 当前局部结构 在锚点单元格内的相对偏移。 */
    rowOffset: number;
    /** 当前局部结构 在锚点单元格内的相对偏移。 */
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
    x: x + emuToPx(anchor.columnOffset),
    y: y + emuToPx(anchor.rowOffset),
  };
}

/** 读取 `readSheetMetrics` 所需的源数据，供 XLSX 解析使用。 */
function readSheetMetrics(sheetNode: Element): SheetMetrics {
  const sheetFormat = childByLocalName(sheetNode, 'sheetFormatPr');
  return {
    defaultColumnWidth: excelWidthToPx(
      Number(
        attr(sheetFormat, 'defaultColWidth') ?? DEFAULT_COLUMN_WIDTH_CHARACTERS,
      ),
    ),
    defaultRowHeight: pointToPx(
      Number(
        attr(sheetFormat, 'defaultRowHeight') ?? DEFAULT_ROW_HEIGHT_POINTS,
      ),
    ),
  };
}

/** 读取 `readColumns` 所需的源数据，供 XLSX 解析使用。 */
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

/** 读取 `readAnchorPoint` 所需的源数据，供 XLSX 解析使用。 */
function readAnchorPoint(node: Element | null) {
  return {
    column: Number(textContent(childByLocalName(node, 'col'))) + 1,
    columnOffset: Number(textContent(childByLocalName(node, 'colOff')) || 0),
    row: Number(textContent(childByLocalName(node, 'row'))) + 1,
    rowOffset: Number(textContent(childByLocalName(node, 'rowOff')) || 0),
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

/** 读取 `readDrawingXml` 所需的源数据，供 XLSX 解析使用。 */
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

/** 读取 `readDrawingBounds` 所需的源数据，供 XLSX 解析使用。 */
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

/** 读取 `readSheetCharts` 所需的源数据，供 XLSX 解析使用。 */
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

/** 读取 `readSheetImages` 所需的源数据，供 XLSX 解析使用。 */
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

/** 读取 `readCellValue` 所需的源数据，供 XLSX 解析使用。 */
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

/** 读取 `readMerges` 所需的源数据，供 XLSX 解析使用。 */
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

/** 读取 `readSheet` 所需的源数据，供 XLSX 解析使用。 */
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
  const metrics = readSheetMetrics(sheetNode);
  const cells = new Map<string, XlsxCell>();
  let maxRow = parsedRange?.endRow ?? 0;
  let maxColumn = parsedRange?.endColumn ?? 0;

  descendantsByLocalName(sheetNode, 'c').forEach((cellNode) => {
    const ref = attr(cellNode, 'r') ?? 'A1';
    const address = parseCellRef(ref);
    const styleId = attr(cellNode, 's')
      ? Number(attr(cellNode, 's'))
      : undefined;
    const value = readCellValue(cellNode, sharedStrings);
    const cell: XlsxCell = {
      ref,
      rowIndex: address.row,
      columnIndex: address.column,
      value: value.value,
      rawValue: value.rawValue,
      type: attr(cellNode, 't'),
      styleId,
      style: resolveStyle(styleId, styleBook),
    };
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

  maxRow = Math.min(Math.max(maxRow, 1), MAX_RENDERED_EMPTY_ROWS);
  maxColumn = Math.min(Math.max(maxColumn, 1), MAX_RENDERED_EMPTY_COLUMNS);

  applyMerges(cells, merges);

  const rowHeights = new Map<number, number>();
  descendantsByLocalName(sheetNode, 'row').forEach((rowNode) => {
    const rowIndex = Number(attr(rowNode, 'r') ?? 0);
    if (rowIndex) {
      rowHeights.set(
        rowIndex,
        pointToPx(Number(attr(rowNode, 'ht')), metrics.defaultRowHeight),
      );
    }
  });
  const columns = readColumns(sheetNode, maxColumn, metrics);

  const rows: XlsxRow[] = Array.from({ length: maxRow }, (_, rowOffset) => {
    const rowIndex = rowOffset + 1;
    return {
      index: rowIndex,
      height: rowHeights.get(rowIndex) ?? metrics.defaultRowHeight,
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
    images: readSheetImages(
      sheetNode,
      sheetInfo.path,
      packageState,
      columns,
      rowHeights,
      metrics,
    ),
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

/** 解析 `parseXlsx` 接收的数据，并返回 XLSX 解析结果。 */
export async function parseXlsx(file: File): Promise<XlsxWorkbook> {
  // sharedStrings 和 styles 是全工作簿共享数据，先解析后再逐个 sheet 套用。
  const entries = await loadXlsxEntries(file);
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
  const sheets = childrenByLocalName(
    childByLocalName(workbookDoc.documentElement, 'sheets'),
    'sheet',
  ).map((sheetNode, index) => {
    const relId = attr(sheetNode, 'r:id') ?? attr(sheetNode, 'id') ?? '';
    const rel = workbookRels[relId];
    const path = rel?.target ?? `xl/worksheets/sheet${index + 1}.xml`;
    return readSheet(
      readXml(entries, path),
      {
        id: attr(sheetNode, 'sheetId') ?? String(index + 1),
        name: decodeMojibake(attr(sheetNode, 'name') ?? `Sheet ${index + 1}`),
        path,
      },
      sharedStrings,
      styleBook,
      packageState,
    );
  });

  return { sheets };
}
