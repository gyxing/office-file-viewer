import {
  resolveOfficeThemeColor,
  type OfficeTheme,
} from '../../shared/ooxml/theme';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  parseXml,
} from '../../shared/ooxml/xml';
import type { SpreadsheetDiagonalBorder } from '../spreadsheet/types';
import type { XlsxCellStyle } from './types';

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
  /** 条件格式 dxfId 引用的差异样式。 */
  differentialStyles: XlsxCellStyle[];
};

/** 完成解析且从 1 开始计数的单元格坐标。 */
export type CellAddress = {
  /** 单元格从 1 开始计数的行索引。 */
  row: number;
  /** 单元格从 1 开始计数的列索引。 */
  column: number;
};

/** 描述从起始单元格到结束单元格的闭合区域。 */
export type CellRange = {
  /** 起始单元格的行索引。 */
  startRow: CellAddress['row'];
  /** 起始单元格的列索引。 */
  startColumn: CellAddress['column'];
  /** 结束单元格的行索引。 */
  endRow: CellAddress['row'];
  /** 结束单元格的列索引。 */
  endColumn: CellAddress['column'];
};

/** Excel 默认行高，单位为磅。 */
export const DEFAULT_ROW_HEIGHT_POINTS = 15;

/** Excel 默认列宽，单位为标准化渲染像素。 */
export const DEFAULT_COLUMN_WIDTH = 64;

/** Excel 默认行高，单位为标准化渲染像素。 */
export const DEFAULT_ROW_HEIGHT = 20;

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

/** 尝试修复由 UTF-8 字节误按单字节编码解读产生的乱码。 */
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

/** 将 Excel 列标签转换为从 1 开始计数的列索引。 */
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

/** 解析 A1 单元格引用并返回从 1 开始计数的坐标。 */
export function parseCellRef(ref: string): CellAddress {
  const match = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return { row: 1, column: 1 };
  return {
    row: Number(match[2]),
    column: columnLabelToIndex(match[1].toUpperCase()),
  };
}

/** 解析 A1 区域引用并返回行列边界。 */
export function parseRange(range?: string): CellRange | undefined {
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

/** 将六位或八位十六进制颜色标准化为 CSS 颜色值。 */
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

/** 按 Excel tint 规则对 RGB 颜色执行明暗调整。 */
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

/** 从无 DOM 的属性映射解析主题色、索引色和 tint。 */
export function parseXlsxColorAttributes(
  attributes: ReadonlyMap<string, string>,
  theme: OfficeTheme,
) {
  if (attributes.get('auto') === '1') return '#000000';
  const rgb = attributes.get('rgb');
  const themeIndex = attributes.get('theme');
  const indexed = attributes.get('indexed');
  const base =
    normalizeHexColor(rgb) ??
    resolveOfficeThemeColor(
      themeIndex ? THEME_COLOR_INDEXES[Number(themeIndex)] : undefined,
      theme,
    ) ??
    normalizeHexColor(indexed ? INDEXED_COLORS[Number(indexed)] : undefined);
  return applyTint(base, attributes.get('tint'));
}

function parseColor(node: Element | null | undefined, theme: OfficeTheme) {
  if (!node) return undefined;
  const attributes = new Map<string, string>();
  ['auto', 'rgb', 'theme', 'indexed', 'tint'].forEach((name) => {
    const value = attr(node, name);
    if (value !== undefined) attributes.set(name, value);
  });
  return parseXlsxColorAttributes(attributes, theme);
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
  if (!xml) {
    return {
      fonts: [],
      fills: [],
      borders: [],
      styles: [],
      differentialStyles: [],
    };
  }
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
      backgroundColor:
        parseColor(childByLocalName(pattern, 'fgColor'), theme) ??
        (attr(pattern, 'patternType') === 'solid'
          ? parseColor(childByLocalName(pattern, 'bgColor'), theme)
          : undefined),
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

  const differentialStyles = childrenByLocalName(
    childByLocalName(styleSheet, 'dxfs'),
    'dxf',
  ).map((dxf): XlsxCellStyle => {
    const font = childByLocalName(dxf, 'font');
    const fill = childByLocalName(dxf, 'fill');
    const pattern = childByLocalName(fill, 'patternFill');
    const border = childByLocalName(dxf, 'border');
    const top = parseBorderStyle(childByLocalName(border, 'top'), theme);
    const right = parseBorderStyle(childByLocalName(border, 'right'), theme);
    const bottom = parseBorderStyle(childByLocalName(border, 'bottom'), theme);
    const left = parseBorderStyle(childByLocalName(border, 'left'), theme);
    return Object.fromEntries(
      Object.entries({
        bold: font ? Boolean(childByLocalName(font, 'b')) : undefined,
        italic: font ? Boolean(childByLocalName(font, 'i')) : undefined,
        underline: font ? Boolean(childByLocalName(font, 'u')) : undefined,
        color: parseColor(childByLocalName(font, 'color'), theme),
        fontSize: pointToCssPx(
          Number(attr(childByLocalName(font, 'sz'), 'val') ?? 0),
        ),
        fontFamily: attr(childByLocalName(font, 'name'), 'val'),
        backgroundColor:
          parseColor(childByLocalName(pattern, 'fgColor'), theme) ??
          (attr(pattern, 'patternType') === 'solid'
            ? parseColor(childByLocalName(pattern, 'bgColor'), theme)
            : undefined),
        border: Boolean(top || right || bottom || left),
        borderTop: borderToCss(top),
        borderRight: borderToCss(right),
        borderBottom: borderToCss(bottom),
        borderLeft: borderToCss(left),
      }).filter(([, value]) => value !== undefined && value !== false),
    ) as XlsxCellStyle;
  });

  return { fonts, fills, borders, styles, differentialStyles };
}

/** 为条件格式解析器开放与普通样式一致的颜色解析。 */
export const parseXlsxColor = parseColor;

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
