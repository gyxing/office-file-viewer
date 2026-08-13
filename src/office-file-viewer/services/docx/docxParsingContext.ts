import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import { readXml } from '../../shared/ooxml/archive';
import {
  collectMedia,
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
  matchesLocalName,
  parseXml,
} from '../../shared/ooxml/xml';
import type { OfficeResourceSource } from '../resource-store/types';
import type { WordBookmarkTarget } from '../word/types';
import {
  readDocxNumbering,
  readDocxNumberingReference,
  type DocxNumberingCatalog,
  type DocxNumberingReference,
} from './parseDocxNumbering';
import type {
  DocxCharacterSpacingControl,
  DocxImage,
  DocxPage,
  DocxTabStop,
  DocxTextStyle,
} from './types';

/** DOCX 缺少页面设置时使用的默认页面尺寸和边距。 */
export const DEFAULT_DOCX_PAGE: DocxPage = {
  width: 794,
  minHeight: 1123,
  marginTop: 96,
  marginRight: 120,
  marginBottom: 96,
  marginLeft: 120,
};

/** 枚举 DOCX 解析可能处于的状态。 */
export type DocxPackageState = {
  /** 压缩包或复合文档包含的条目。 */
  entries: OfficeEntryMap;
  /** 按关系标识索引的 OOXML 关系。 */
  relationships: Record<string, Record<string, OfficeRelationship>>;
  /** 按 OOXML 包内路径索引的媒体资源映射。 */
  mediaByPath: Record<string, OfficeResourceSource>;
  /** 按媒体文件名索引的资源映射。 */
  mediaByName: Record<string, OfficeResourceSource>;
};

/** 汇总 DOCX 解析各步骤共享的上下文。 */
export type DocxParseContext = {
  /** 解析各 DOCX 部件时共享的包状态。 */
  packageState: DocxPackageState;
  /** 按关系标识索引的 DOCX 主文档关系。 */
  documentRels: Record<string, OfficeRelationship>;
  /** 文档使用的主题颜色和字体配置。 */
  theme: OfficeTheme;
  /** 文档网格推导出的默认正文行高，单位为标准化渲染像素。 */
  defaultLineHeight?: number;
  /** 文档网格单行的实际高度，供表格等局部排版遵循吸附规则。 */
  documentGridLineHeight?: number;
  /** 源文档声明的东亚标点字符间距压缩方式。 */
  characterSpacingControl?: DocxCharacterSpacingControl;
  /** DOCX 自动编号定义及当前解析计数状态。 */
  numbering: DocxNumberingCatalog;
  /** 按样式标识索引的 DOCX 样式目录。 */
  styles: DocxStyleCatalog;
  /** 解析过程中已收集的图片资源。 */
  images: DocxImage[];
  /** 解析过程中按源名称收集的书签定位信息。 */
  bookmarks: Record<string, WordBookmarkTarget>;
  /** 下一张图片使用的零基索引。 */
  imageIndex: number;
  /** 图表在所属图表集合中的索引。 */
  chartIndex: number;
  /** 下一个形状使用的零基索引。 */
  shapeIndex: number;
};

/** 读取 DOCX 容器节点下块级内容时使用的选项。 */
export type ReadBlockChildrenOptions = {
  /** 是否位于形状内部；用于选择局部坐标和排版规则。 */
  insideShape?: boolean;
  /** 是否位于表格单元格内部。 */
  insideTable?: boolean;
  /** 是否位于页眉或页脚区域；这些段落不属于正文大纲。 */
  insidePageRegion?: boolean;
};

/** DOCX 单个段落或字符样式的继承和排版属性。 */
export type DocxStyleDefinition = {
  /** 样式适用的对象类型。 */
  kind: 'paragraph' | 'character' | 'table';
  /** 源 styles.xml 中声明的样式名称。 */
  name?: string;
  /** 当前样式继承的父样式标识。 */
  basedOn?: string;
  /** 样式直接声明的大纲级别，使用从 0 开始的内部表示。 */
  outlineLevel?: number;
  /** 样式直接声明的自动编号引用。 */
  numbering?: DocxNumberingReference;
  /** 样式直接声明的制表位。 */
  tabStops?: DocxTabStop[];
  /** 样式是否要求文字吸附文档网格。 */
  snapToGrid?: boolean;
  /** 当前内容使用的渲染样式。 */
  style: DocxTextStyle;
};

/** 按样式标识索引的 DOCX 样式目录。 */
export type DocxStyleCatalog = {
  /** 默认段落、文本与样式标识配置。 */
  defaults: {
    /** 默认段落样式。 */
    paragraph?: DocxTextStyle;
    /** 默认文字样式。 */
    run?: DocxTextStyle;
    /** 默认段落样式标识。 */
    paragraphStyleId?: string;
    /** 默认表格样式标识。 */
    tableStyleId?: string;
  };
  /** 按业务键索引的 映射。 */
  styles: Record<string, DocxStyleDefinition>;
  /** 按样式缓存继承后的大纲级别，null 表示没有大纲语义。 */
  outlineLevelCache: Map<string, number | null>;
  /** 按样式缓存 TOC 判定，避免为每个段落重复遍历继承链。 */
  tocStyleCache: Map<string, boolean>;
};

/** DOCX 缺少字体信息时使用的默认字体回退栈。 */
const DEFAULT_DOCX_FONT_FAMILY =
  '"Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", Arial, sans-serif';

// DOCX 与 PPTX 类似是 zip 包结构，正文、样式、主题、媒体通过关系文件互相引用。
/** 根据已读取条目和可选懒媒体索引建立 DOCX 包上下文。 */
export function buildDocxPackageState(
  entries: OfficeEntryMap,
  lazyMedia?: {
    byPath: Record<string, OfficeResourceSource>;
    byName: Record<string, OfficeResourceSource>;
  },
): DocxPackageState {
  const relationships: DocxPackageState['relationships'] = {};

  for (const [path, value] of entries) {
    if (typeof value === 'string' && path.endsWith('.rels')) {
      relationships[path] = readRelationships(value, path);
    }
  }

  const materializedMedia = lazyMedia
    ? undefined
    : collectMedia(entries, 'word/media/');
  const media = lazyMedia ?? {
    byPath: Object.fromEntries(
      Object.entries(materializedMedia?.byPath ?? {}).map(([path, url]) => [
        path,
        { kind: 'url' as const, url },
      ]),
    ),
    byName: Object.fromEntries(
      Object.entries(materializedMedia?.byName ?? {}).map(([name, url]) => [
        name,
        { kind: 'url' as const, url },
      ]),
    ),
  };

  return {
    entries,
    relationships,
    mediaByPath: media.byPath,
    mediaByName: media.byName,
  };
}

export function twipToPx(value?: string | number) {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue)) return undefined;
  return (numberValue / 1440) * 96;
}

function pointToPx(value?: string | number) {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue)) return undefined;
  return numberValue * (96 / 72);
}

export function positiveTwipToPx(value?: string | number) {
  const result = twipToPx(value);
  return result !== undefined && result >= 0 ? result : undefined;
}

function eighthPointToPx(value?: string | number) {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue) || numberValue <= 0) return undefined;
  return (numberValue / 8) * (96 / 72);
}

export function vmlUnitToPx(value?: string | number) {
  const raw = String(value ?? '').trim();
  if (!raw) return undefined;
  const match = raw.match(/^(-?\d+(?:\.\d+)?)(pt|px|in|cm|mm)?$/i);
  if (!match) return undefined;
  const numberValue = Number(match[1]);
  if (!Number.isFinite(numberValue)) return undefined;
  const unit = match[2]?.toLowerCase();
  if (unit === 'px') return numberValue;
  if (unit === 'in') return numberValue * 96;
  if (unit === 'cm') return (numberValue / 2.54) * 96;
  if (unit === 'mm') return (numberValue / 25.4) * 96;
  if (unit === 'pt') return pointToPx(numberValue);
  return emuToPx(numberValue);
}

export function readCssDeclaration(style: string | undefined, name: string) {
  if (!style) return undefined;
  const match = style.match(
    new RegExp(`(?:^|;)\\s*${name}\\s*:\\s*([^;]+)`, 'i'),
  );
  return match?.[1]?.trim();
}

export function readCssSize(
  style: string | undefined,
  name: string,
  scale?: number,
) {
  const raw = readCssDeclaration(style, name);
  if (!raw) return undefined;
  if (scale !== undefined && /^-?\d+(?:\.\d+)?$/.test(raw)) {
    return Number(raw) * scale;
  }
  return vmlUnitToPx(raw);
}

export function readCssPosition(
  style: string | undefined,
  name: 'left' | 'top',
) {
  return readCssSize(style, `margin-${name}`) ?? readCssSize(style, name);
}

function readDocxLineHeight(spacingNode: Element | null | undefined) {
  const value = Number(
    attr(spacingNode, 'w:line') ?? attr(spacingNode, 'line'),
  );
  if (!Number.isFinite(value) || value <= 0) return undefined;
  const rule = attr(spacingNode, 'w:lineRule') ?? attr(spacingNode, 'lineRule');
  if (rule === 'exact' || rule === 'atLeast') {
    return twipToPx(value);
  }
  return value / 240;
}

/** 读取 OOXML 行距规则，避免精确行距被正文网格扩张。 */
function readDocxLineHeightRule(
  spacingNode: Element | null | undefined,
): DocxTextStyle['lineHeightRule'] {
  const value = Number(
    attr(spacingNode, 'w:line') ?? attr(spacingNode, 'line'),
  );
  if (!Number.isFinite(value) || value <= 0) return undefined;
  const rule = attr(spacingNode, 'w:lineRule') ?? attr(spacingNode, 'lineRule');
  return rule === 'exact' || rule === 'atLeast' ? rule : 'auto';
}

function halfPointToPx(value?: string) {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue)) return undefined;
  return (numberValue / 2) * (96 / 72);
}

export function parseHexColor(value?: string) {
  if (!value || value === 'auto' || value === 'none') return undefined;
  // 提取颜色值，忽略额外的信息（如 "#41719C [3204]"）
  const match = value.match(/^#?([0-9a-f]{6})/i);
  if (!match) return undefined;
  return `#${match[1]}`;
}

/** 将输入标准化为 `normalizeCssColor` 返回的结构。 */
export function normalizeCssColor(value?: string) {
  if (!value || value === 'auto' || value === 'none') return undefined;
  return value.startsWith('#') ? value : parseHexColor(value);
}

export function readVal(node: Element | null | undefined) {
  return attr(node, 'w:val') ?? attr(node, 'val');
}

/** 将 Word 对齐值映射为统一文本样式。 */
export function mapAlignment(
  value?: string,
): DocxTextStyle['align'] | undefined {
  if (
    value === 'left' ||
    value === 'center' ||
    value === 'right' ||
    value === 'justify'
  )
    return value;
  if (value === 'both') return 'justify';
  return undefined;
}

/** Word 高亮颜色名称到 CSS 颜色值的映射。 */
const WORD_HIGHLIGHT_COLORS: Record<string, string> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00ffff',
  green: '#00ff00',
  magenta: '#ff00ff',
  red: '#ff0000',
  yellow: '#ffff00',
  white: '#ffffff',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#c0c0c0',
};

export function clamp255(value: number) {
  return Math.max(0, Math.min(255, Math.round(value)));
}

/** 应用 DrawingML 颜色变换，避免浅色主题填充退化为高饱和度基色。 */
function applyDrawingColorTransforms(
  color: string | undefined,
  colorNode: Element | null | undefined,
) {
  if (!color || !colorNode) return color;
  const normalized = color.replace('#', '');
  if (!/^[0-9a-f]{6}$/i.test(normalized)) return color;

  const rgb = Number.parseInt(normalized, 16);
  let red = (rgb >> 16) & 255;
  let green = (rgb >> 8) & 255;
  let blue = rgb & 255;
  Array.from(colorNode.children).forEach((transform) => {
    const value = Number(attr(transform, 'val'));
    if (!Number.isFinite(value)) return;
    const ratio = Math.max(0, Math.min(1, value / 100000));
    if (matchesLocalName(transform, 'tint')) {
      red += (255 - red) * ratio;
      green += (255 - green) * ratio;
      blue += (255 - blue) * ratio;
    } else if (matchesLocalName(transform, 'shade')) {
      red *= ratio;
      green *= ratio;
      blue *= ratio;
    } else if (matchesLocalName(transform, 'lumMod')) {
      red *= ratio;
      green *= ratio;
      blue *= ratio;
    } else if (matchesLocalName(transform, 'lumOff')) {
      red += 255 * ratio;
      green += 255 * ratio;
      blue += 255 * ratio;
    }
  });

  return `#${[red, green, blue]
    .map((channel) => clamp255(channel).toString(16).padStart(2, '0'))
    .join('')}`;
}

/** 读取 DrawingML 直接色、主题色及其明暗变换。 */
export function readDrawingColor(
  node: Element | null | undefined,
  theme: OfficeTheme,
) {
  if (!node) return undefined;
  const solidFill =
    childByLocalName(node, 'solidFill') ??
    (matchesLocalName(node, 'solidFill') ? node : null);
  if (!solidFill) return undefined;
  const srgb = childByLocalName(solidFill, 'srgbClr');
  const scheme = childByLocalName(solidFill, 'schemeClr');
  const sys = childByLocalName(solidFill, 'sysClr');
  const isTransparent = (colorNode: Element | null | undefined) =>
    attr(childByLocalName(colorNode, 'alpha'), 'val') === '0';
  return (
    (isTransparent(srgb)
      ? undefined
      : applyDrawingColorTransforms(parseHexColor(attr(srgb, 'val')), srgb)) ??
    (isTransparent(scheme)
      ? undefined
      : applyDrawingColorTransforms(
          resolveOfficeThemeColor(attr(scheme, 'val'), theme),
          scheme,
        )) ??
    (isTransparent(sys)
      ? undefined
      : applyDrawingColorTransforms(
          parseHexColor(attr(sys, 'lastClr') ?? attr(sys, 'val')),
          sys,
        ))
  );
}

function tintHexColor(color: string | undefined, tint?: string) {
  if (!color || !tint) return color;
  const normalized = color.replace('#', '');
  if (!/^[0-9a-f]{6}$/i.test(normalized)) return color;
  const tintValue = Number.parseInt(tint, 16);
  if (!Number.isFinite(tintValue)) return color;
  const ratio = Math.max(0, Math.min(1, tintValue / 255));
  const rgb = Number.parseInt(normalized, 16);
  const r = (rgb >> 16) & 255;
  const g = (rgb >> 8) & 255;
  const b = rgb & 255;
  return `#${[r, g, b]
    .map((channel) =>
      clamp255(channel + (255 - channel) * ratio)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`;
}

function shadeHexColor(color: string | undefined, shade?: string) {
  if (!color || !shade) return color;
  const normalized = color.replace('#', '');
  if (!/^[0-9a-f]{6}$/i.test(normalized)) return color;
  const shadeValue = Number.parseInt(shade, 16);
  if (!Number.isFinite(shadeValue)) return color;
  const ratio = Math.max(0, Math.min(1, shadeValue / 255));
  const rgb = Number.parseInt(normalized, 16);
  const r = (rgb >> 16) & 255;
  const g = (rgb >> 8) & 255;
  const b = rgb & 255;
  return `#${[r, g, b]
    .map((channel) =>
      clamp255(channel * ratio)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`;
}

/** 解析并确定 `resolveThemeFillColor` 对应的引用或配置。 */
function resolveThemeFillColor(
  node: Element | null | undefined,
  theme: OfficeTheme,
) {
  const themeFill = attr(node, 'w:themeFill') ?? attr(node, 'themeFill');
  const themeColor = resolveOfficeThemeColor(themeFill, theme);
  return shadeHexColor(
    tintHexColor(
      themeColor,
      attr(node, 'w:themeFillTint') ?? attr(node, 'themeFillTint'),
    ),
    attr(node, 'w:themeFillShade') ?? attr(node, 'themeFillShade'),
  );
}

export function readShading(
  node: Element | null | undefined,
  theme: OfficeTheme,
) {
  if (!node) return undefined;
  const directFill = parseHexColor(attr(node, 'w:fill') ?? attr(node, 'fill'));
  return directFill ?? resolveThemeFillColor(node, theme);
}

function readHighlight(node: Element | null | undefined) {
  const value = readVal(node);
  if (!value || value === 'none') return undefined;
  return WORD_HIGHLIGHT_COLORS[value] ?? parseHexColor(value);
}

export function readBorder(node: Element | null | undefined) {
  const value = readVal(node);
  if (!node || !value || value === 'none' || value === 'nil') return undefined;
  const color =
    parseHexColor(attr(node, 'w:color') ?? attr(node, 'color')) ?? '#000';
  const width = eighthPointToPx(attr(node, 'w:sz') ?? attr(node, 'sz')) ?? 1;
  const style =
    value === 'dashed' || value === 'dashSmallGap'
      ? 'dashed'
      : value === 'dotted'
      ? 'dotted'
      : 'solid';
  return `${width}px ${style} ${color}`;
}

function readParagraphBorders(pPr: Element | null | undefined) {
  const pBdr = childByLocalName(pPr, 'pBdr');
  return {
    borderTop: readBorder(childByLocalName(pBdr, 'top')),
    borderRight: readBorder(childByLocalName(pBdr, 'right')),
    borderBottom: readBorder(childByLocalName(pBdr, 'bottom')),
    borderLeft: readBorder(childByLocalName(pBdr, 'left')),
    paddingTop: pointToPx(
      attr(childByLocalName(pBdr, 'top'), 'w:space') ??
        attr(childByLocalName(pBdr, 'top'), 'space'),
    ),
    paddingRight: pointToPx(
      attr(childByLocalName(pBdr, 'right'), 'w:space') ??
        attr(childByLocalName(pBdr, 'right'), 'space'),
    ),
    paddingBottom: pointToPx(
      attr(childByLocalName(pBdr, 'bottom'), 'w:space') ??
        attr(childByLocalName(pBdr, 'bottom'), 'space'),
    ),
    paddingLeft: pointToPx(
      attr(childByLocalName(pBdr, 'left'), 'w:space') ??
        attr(childByLocalName(pBdr, 'left'), 'space'),
    ),
  };
}

function readParagraphPropertyStyle(
  pPr: Element | null | undefined,
  theme: OfficeTheme,
): DocxTextStyle | undefined {
  if (!pPr) return undefined;
  const spacing = childByLocalName(pPr, 'spacing');
  const ind = childByLocalName(pPr, 'ind');
  const style: DocxTextStyle = {
    align: mapAlignment(readVal(childByLocalName(pPr, 'jc'))),
    spacingBefore: positiveTwipToPx(
      attr(spacing, 'w:before') ?? attr(spacing, 'before'),
    ),
    spacingAfter: positiveTwipToPx(
      attr(spacing, 'w:after') ?? attr(spacing, 'after'),
    ),
    indentLeft: twipToPx(attr(ind, 'w:left') ?? attr(ind, 'left')),
    indentRight: twipToPx(attr(ind, 'w:right') ?? attr(ind, 'right')),
    firstLineIndent: twipToPx(
      attr(ind, 'w:firstLine') ?? attr(ind, 'firstLine'),
    ),
    lineHeight: readDocxLineHeight(spacing),
    lineHeightRule: readDocxLineHeightRule(spacing),
    backgroundColor: readShading(childByLocalName(pPr, 'shd'), theme),
    ...readParagraphBorders(pPr),
  };
  const cleaned = Object.fromEntries(
    Object.entries(style).filter(([, value]) => value !== undefined),
  ) as DocxTextStyle;
  return Object.keys(cleaned).length ? cleaned : undefined;
}

/** 读取段落制表位，保留目录的右对齐位置与引导符语义。 */
function readParagraphTabStops(
  pPr: Element | null | undefined,
): DocxTabStop[] | undefined {
  const tabs = childByLocalName(pPr, 'tabs');
  if (!tabs) return undefined;
  const stops = childrenByLocalName(tabs, 'tab')
    .map((tab): DocxTabStop | undefined => {
      const rawAlign = readVal(tab);
      const position = twipToPx(attr(tab, 'w:pos') ?? attr(tab, 'pos'));
      if (position === undefined || rawAlign === 'clear') return undefined;
      const align =
        rawAlign === 'center' ||
        rawAlign === 'right' ||
        rawAlign === 'decimal' ||
        rawAlign === 'bar' ||
        rawAlign === 'num'
          ? rawAlign === 'num'
            ? 'number'
            : rawAlign
          : 'left';
      const rawLeader = attr(tab, 'w:leader') ?? attr(tab, 'leader');
      const leader =
        rawLeader === 'dot' ||
        rawLeader === 'hyphen' ||
        rawLeader === 'underscore' ||
        rawLeader === 'middleDot' ||
        rawLeader === 'none'
          ? rawLeader
          : undefined;
      return { position, align, leader };
    })
    .filter((stop): stop is DocxTabStop => Boolean(stop));
  return stops.length ? stops : undefined;
}

export function readOnOff(node: Element | null | undefined) {
  if (!node) return undefined;
  const value = attr(node, 'w:val') ?? attr(node, 'val');
  if (value === undefined) return true;
  return value !== '0' && value !== 'false' && value !== 'off';
}

function firstDefined<T>(...values: Array<T | undefined>) {
  return values.find((value) => value !== undefined);
}

function readThemeFont(
  rPr: Element | null | undefined,
  rFonts: Element | null | undefined,
  theme: OfficeTheme,
) {
  const themeFont =
    attr(rFonts, 'w:eastAsiaTheme') ??
    attr(rFonts, 'eastAsiaTheme') ??
    attr(rFonts, 'w:asciiTheme') ??
    attr(rFonts, 'asciiTheme') ??
    attr(rFonts, 'w:hAnsiTheme') ??
    attr(rFonts, 'hAnsiTheme') ??
    attr(rFonts, 'w:cstheme') ??
    attr(rFonts, 'cstheme');
  if (!themeFont) return undefined;

  const bucket = themeFont.toLowerCase().includes('major')
    ? 'majorFont'
    : 'minorFont';
  const language = (
    attr(childByLocalName(rPr, 'lang'), 'w:eastAsia') ??
    attr(childByLocalName(rPr, 'lang'), 'eastAsia') ??
    ''
  ).toLowerCase();
  const script = language.startsWith('ja')
    ? 'Jpan'
    : language.startsWith('ko')
    ? 'Hang'
    : /zh-(tw|hk|mo)/.test(language)
    ? 'Hant'
    : language.startsWith('zh')
    ? 'Hans'
    : undefined;
  return (
    (script ? theme.fontScheme?.[`${bucket}:${script}`] : undefined) ??
    theme.fontScheme?.[bucket]
  );
}

function quoteFontFamily(value?: string) {
  if (!value) return undefined;
  return value
    .split(',')
    .map((font) => font.trim())
    .filter(Boolean)
    .map((font) =>
      /^["'].*["']$/.test(font) || /^[a-z-]+$/i.test(font) ? font : `"${font}"`,
    )
    .join(', ');
}

function readFontFamily(
  rPr: Element | null | undefined,
  theme: OfficeTheme,
  allowFallback = false,
) {
  const rFonts = childByLocalName(rPr, 'rFonts');
  const ascii = attr(rFonts, 'w:ascii') ?? attr(rFonts, 'ascii');
  const eastAsia = attr(rFonts, 'w:eastAsia') ?? attr(rFonts, 'eastAsia');
  const hAnsi = attr(rFonts, 'w:hAnsi') ?? attr(rFonts, 'hAnsi');
  const cs = attr(rFonts, 'w:cs') ?? attr(rFonts, 'cs');
  const themeFonts = theme.fontScheme ?? {};
  const themeFont = readThemeFont(rPr, rFonts, theme);
  const scriptHint = attr(rFonts, 'w:hint') ?? attr(rFonts, 'hint');
  // 字体脚本提示决定 CSS 字体链优先级，避免系统字体链接为中文粗体选择错误的替代字体。
  const fontCandidates =
    scriptHint === 'eastAsia'
      ? [eastAsia, ascii ?? hAnsi, themeFont, cs]
      : scriptHint === 'cs'
      ? [cs, ascii ?? hAnsi, eastAsia, themeFont]
      : [ascii ?? hAnsi, eastAsia, themeFont, cs];
  const explicitFonts = fontCandidates
    .filter((font): font is string => Boolean(font))
    .filter((font, index, fonts) => fonts.indexOf(font) === index);
  if (explicitFonts.length) return quoteFontFamily(explicitFonts.join(','));
  if (!allowFallback) return undefined;
  return quoteFontFamily(
    themeFonts.minorFont ?? themeFonts.majorFont ?? DEFAULT_DOCX_FONT_FAMILY,
  );
}

function readDocxStyles(
  entries: OfficeEntryMap,
  theme: OfficeTheme,
): DocxStyleCatalog {
  // styles.xml 会提供默认样式和命名样式，段落/文字解析时再与直接格式合并。
  const xml = readXml(entries, 'word/styles.xml');
  if (!xml)
    return {
      defaults: {},
      styles: {},
      outlineLevelCache: new Map(),
      tocStyleCache: new Map(),
    };

  const doc = parseXml(xml);
  const root = doc.documentElement;
  const styles: Record<string, DocxStyleDefinition> = {};
  const defaults: DocxStyleCatalog['defaults'] = {};

  const docDefaults = childByLocalName(root, 'docDefaults');
  const rPrDefault = childByLocalName(
    childByLocalName(docDefaults, 'rPrDefault'),
    'rPr',
  );
  const pPrDefault = childByLocalName(
    childByLocalName(docDefaults, 'pPrDefault'),
    'pPr',
  );
  defaults.run = readTextStyle(rPrDefault, theme, true);
  defaults.paragraph = mergeTextStyle(
    readParagraphPropertyStyle(pPrDefault, theme),
    readTextStyle(childByLocalName(pPrDefault, 'rPr'), theme, true),
  );

  childrenByLocalName(root, 'style').forEach((styleNode) => {
    const styleId = attr(styleNode, 'styleId');
    const kindAttr = attr(styleNode, 'type');
    if (!styleId) return;
    if (kindAttr === 'paragraph' && attr(styleNode, 'w:default') === '1')
      defaults.paragraphStyleId = styleId;
    if (kindAttr === 'table' && attr(styleNode, 'w:default') === '1')
      defaults.tableStyleId = styleId;

    const basedOn =
      attr(childByLocalName(styleNode, 'basedOn'), 'w:val') ??
      attr(childByLocalName(styleNode, 'basedOn'), 'val') ??
      undefined;
    const name =
      attr(childByLocalName(styleNode, 'name'), 'w:val') ??
      attr(childByLocalName(styleNode, 'name'), 'val') ??
      styleId;
    const pPr = childByLocalName(styleNode, 'pPr');
    const rPr = childByLocalName(styleNode, 'rPr');

    let style: DocxTextStyle | undefined;
    if (kindAttr === 'paragraph') {
      style = mergeTextStyle(
        readParagraphPropertyStyle(pPr, theme),
        readTextStyle(rPr, theme),
      );
    } else if (kindAttr === 'table') {
      const tblPr = childByLocalName(styleNode, 'tblPr');
      style = readParagraphPropertyStyle(childByLocalName(tblPr, 'pPr'), theme);
    } else {
      style = readTextStyle(rPr, theme);
    }

    styles[styleId] = {
      kind:
        kindAttr === 'paragraph'
          ? 'paragraph'
          : kindAttr === 'table'
          ? 'table'
          : 'character',
      name,
      basedOn,
      outlineLevel: readDocxOutlineLevel(pPr),
      numbering: readDocxNumberingReference(pPr),
      tabStops: readParagraphTabStops(pPr),
      snapToGrid: readOnOff(childByLocalName(pPr, 'snapToGrid')),
      style: style ?? {},
    };
  });

  return {
    defaults,
    styles,
    outlineLevelCache: new Map(),
    tocStyleCache: new Map(),
  };
}

function readUnderline(rPr: Element | null | undefined) {
  const underline = childByLocalName(rPr, 'u');
  if (!underline) return undefined;
  const value = attr(underline, 'w:val') ?? attr(underline, 'val');
  return value !== 'none' && value !== '0' && value !== 'false';
}

function readTextStyle(
  rPr: Element | null | undefined,
  theme: OfficeTheme,
  allowFontFallback = false,
): DocxTextStyle | undefined {
  if (!rPr) return undefined;

  const runFonts = childByLocalName(rPr, 'rFonts');
  const scriptHint =
    attr(runFonts, 'w:hint') ?? attr(runFonts, 'hint') ?? undefined;
  const fontHint =
    scriptHint === 'default' || scriptHint === 'eastAsia' || scriptHint === 'cs'
      ? scriptHint
      : undefined;
  const usesComplexScript =
    fontHint === 'cs' || Boolean(childByLocalName(rPr, 'cs'));
  const color =
    readDrawingColor(childByLocalName(rPr, 'textFill'), theme) ??
    parseHexColor(
      attr(childByLocalName(rPr, 'color'), 'w:val') ??
        attr(childByLocalName(rPr, 'color'), 'val'),
    );
  const fontSize = halfPointToPx(
    attr(childByLocalName(rPr, 'sz'), 'w:val') ??
      attr(childByLocalName(rPr, 'sz'), 'val'),
  );
  const complexScriptFontSize = halfPointToPx(
    attr(childByLocalName(rPr, 'szCs'), 'w:val') ??
      attr(childByLocalName(rPr, 'szCs'), 'val'),
  );
  const fontFamily = readFontFamily(rPr, theme, allowFontFallback);
  // szCs 只有与复杂脚本或显式字体声明成对出现时才参与行盒，单独的兼容字号不能撑高普通东亚文字。
  const lineBoxFontSize =
    complexScriptFontSize !== undefined && (usesComplexScript || fontFamily)
      ? Math.max(fontSize ?? 0, complexScriptFontSize)
      : fontSize;
  const style: DocxTextStyle = {
    bold: firstDefined(
      readOnOff(childByLocalName(rPr, 'b')),
      usesComplexScript ? readOnOff(childByLocalName(rPr, 'bCs')) : undefined,
    ),
    italic: firstDefined(
      readOnOff(childByLocalName(rPr, 'i')),
      usesComplexScript ? readOnOff(childByLocalName(rPr, 'iCs')) : undefined,
    ),
    underline: readUnderline(rPr),
    strike: firstDefined(
      readOnOff(childByLocalName(rPr, 'strike')),
      readOnOff(childByLocalName(rPr, 'dstrike')),
    ),
    smallCaps: readOnOff(childByLocalName(rPr, 'smallCaps')),
    allCaps: readOnOff(childByLocalName(rPr, 'caps')),
    color,
    backgroundColor:
      readHighlight(childByLocalName(rPr, 'highlight')) ??
      readShading(childByLocalName(rPr, 'shd'), theme),
    fontHint,
    fontSize,
    letterSpacing: twipToPx(
      attr(childByLocalName(rPr, 'spacing'), 'w:val') ??
        attr(childByLocalName(rPr, 'spacing'), 'val'),
    ),
    lineBoxFontSize: lineBoxFontSize || undefined,
    fontFamily,
  };

  const cleaned = Object.fromEntries(
    Object.entries(style).filter(([, value]) => value !== undefined),
  ) as DocxTextStyle;
  return Object.keys(cleaned).length ? cleaned : undefined;
}

/** 合并 `mergeTwoTextStyles` 接收的多份数据。 */
function mergeTwoTextStyles(
  base?: DocxTextStyle,
  next?: DocxTextStyle,
): DocxTextStyle {
  return {
    ...base,
    ...next,
    fontSize: next?.fontSize ?? base?.fontSize,
    lineBoxFontSize: next?.lineBoxFontSize ?? base?.lineBoxFontSize,
    fontFamily: next?.fontFamily ?? base?.fontFamily,
    color: next?.color ?? base?.color,
    lineHeight: next?.lineHeight ?? base?.lineHeight,
    lineHeightRule: next?.lineHeightRule ?? base?.lineHeightRule,
    spacingBefore: next?.spacingBefore ?? base?.spacingBefore,
    spacingAfter: next?.spacingAfter ?? base?.spacingAfter,
    indentLeft: next?.indentLeft ?? base?.indentLeft,
    indentRight: next?.indentRight ?? base?.indentRight,
    firstLineIndent: next?.firstLineIndent ?? base?.firstLineIndent,
    backgroundColor: next?.backgroundColor ?? base?.backgroundColor,
    borderTop: next?.borderTop ?? base?.borderTop,
    borderRight: next?.borderRight ?? base?.borderRight,
    borderBottom: next?.borderBottom ?? base?.borderBottom,
    borderLeft: next?.borderLeft ?? base?.borderLeft,
    paddingTop: next?.paddingTop ?? base?.paddingTop,
    paddingRight: next?.paddingRight ?? base?.paddingRight,
    paddingBottom: next?.paddingBottom ?? base?.paddingBottom,
    paddingLeft: next?.paddingLeft ?? base?.paddingLeft,
    align: next?.align ?? base?.align,
  };
}

/** 合并 `mergeTextStyle` 接收的多份数据。 */
export function mergeTextStyle(
  ...styles: Array<DocxTextStyle | undefined>
): DocxTextStyle | undefined {
  const merged = styles.reduce<DocxTextStyle>(
    (acc, style) => mergeTwoTextStyles(acc, style),
    {},
  );
  return Object.keys(merged).length ? merged : undefined;
}

/** 解析并确定 `resolveDocxStyle` 对应的引用或配置。 */
export function resolveDocxStyle(
  styleId: string | undefined,
  catalog: DocxStyleCatalog,
  seen: Set<string> = new Set(),
): DocxTextStyle | undefined {
  if (!styleId || seen.has(styleId)) return undefined;
  const entry = catalog.styles[styleId];
  if (!entry) return undefined;
  seen.add(styleId);
  const base = resolveDocxStyle(entry.basedOn, catalog, seen);
  return mergeTextStyle(base, entry.style);
}

/** 读取段落属性直接声明的 OOXML 大纲级别。 */
function readDocxOutlineLevel(
  pPr: Element | null | undefined,
): number | undefined {
  const raw =
    attr(childByLocalName(pPr, 'outlineLvl'), 'w:val') ??
    attr(childByLocalName(pPr, 'outlineLvl'), 'val');
  if (raw === undefined) return undefined;
  const level = Number.parseInt(raw, 10);
  return Number.isInteger(level) && level >= 0 && level <= 9
    ? level
    : undefined;
}

/** 沿 basedOn 继承链解析样式大纲级别，并缓存结果。 */
function resolveDocxOutlineLevel(
  styleId: string | undefined,
  catalog: DocxStyleCatalog,
  seen: Set<string> = new Set(),
): number | undefined {
  if (!styleId) return undefined;
  if (catalog.outlineLevelCache.has(styleId))
    return catalog.outlineLevelCache.get(styleId) ?? undefined;
  if (seen.has(styleId)) {
    catalog.outlineLevelCache.set(styleId, null);
    return undefined;
  }

  const entry = catalog.styles[styleId];
  if (!entry || entry.kind !== 'paragraph') {
    catalog.outlineLevelCache.set(styleId, null);
    return undefined;
  }
  seen.add(styleId);
  const level =
    entry.outlineLevel ?? resolveDocxOutlineLevel(entry.basedOn, catalog, seen);
  catalog.outlineLevelCache.set(styleId, level ?? null);
  return level;
}

/** 判断段落样式是否为目录域生成的 TOC 样式。 */
function isDocxTocStyle(
  styleId: string | undefined,
  catalog: DocxStyleCatalog,
  seen: Set<string> = new Set(),
): boolean {
  if (!styleId) return false;
  const cached = catalog.tocStyleCache.get(styleId);
  if (cached !== undefined) return cached;
  if (seen.has(styleId)) {
    catalog.tocStyleCache.set(styleId, false);
    return false;
  }

  const entry = catalog.styles[styleId];
  if (!entry) {
    catalog.tocStyleCache.set(styleId, false);
    return false;
  }
  seen.add(styleId);
  const isToc =
    /^toc\s*\d+$/i.test(entry.name ?? styleId) ||
    isDocxTocStyle(entry.basedOn, catalog, seen);
  catalog.tocStyleCache.set(styleId, isToc);
  return isToc;
}

/** 沿 basedOn 继承链读取段落样式中的非文字属性。 */
function resolveDocxParagraphStyleProperty<
  K extends 'numbering' | 'tabStops' | 'snapToGrid',
>(
  styleId: string | undefined,
  catalog: DocxStyleCatalog,
  property: K,
  seen: Set<string> = new Set(),
): DocxStyleDefinition[K] | undefined {
  if (!styleId || seen.has(styleId)) return undefined;
  const entry = catalog.styles[styleId];
  if (!entry || entry.kind !== 'paragraph') return undefined;
  seen.add(styleId);
  return (
    entry[property] ??
    resolveDocxParagraphStyleProperty(entry.basedOn, catalog, property, seen)
  );
}

/** 解析并确定 `resolveParagraphStyle` 对应的引用或配置。 */
export function resolveParagraphStyle(
  pPr: Element | null | undefined,
  catalog: DocxStyleCatalog,
  theme: OfficeTheme,
) {
  // 段落正文继承命名样式；pPr/rPr 只描述段落标记，不能覆盖可见 run 的文字格式。
  const styleId =
    attr(childByLocalName(pPr, 'pStyle'), 'w:val') ??
    attr(childByLocalName(pPr, 'pStyle'), 'val');
  const effectiveStyleId = styleId ?? catalog.defaults.paragraphStyleId;
  const baseStyle = resolveDocxStyle(
    catalog.defaults.paragraphStyleId,
    catalog,
  );
  const namedStyle = resolveDocxStyle(styleId, catalog);
  const style = mergeTextStyle(
    catalog.defaults.paragraph,
    baseStyle,
    namedStyle,
  );
  const directStyle = readParagraphPropertyStyle(pPr, theme);
  const contentStyle = mergeTextStyle(style, directStyle);
  const paragraphMarkStyle = mergeTextStyle(
    contentStyle,
    readTextStyle(childByLocalName(pPr, 'rPr'), theme),
  );
  return {
    align: directStyle?.align ?? style?.align,
    spacingBefore: directStyle?.spacingBefore ?? style?.spacingBefore,
    spacingAfter: directStyle?.spacingAfter ?? style?.spacingAfter,
    indentLeft: directStyle?.indentLeft ?? style?.indentLeft,
    indentRight: directStyle?.indentRight ?? style?.indentRight,
    firstLineIndent: directStyle?.firstLineIndent ?? style?.firstLineIndent,
    lineHeight: directStyle?.lineHeight ?? style?.lineHeight,
    lineHeightRule:
      directStyle?.lineHeightRule ?? style?.lineHeightRule,
    backgroundColor: directStyle?.backgroundColor ?? style?.backgroundColor,
    borderTop: directStyle?.borderTop ?? style?.borderTop,
    borderRight: directStyle?.borderRight ?? style?.borderRight,
    borderBottom: directStyle?.borderBottom ?? style?.borderBottom,
    borderLeft: directStyle?.borderLeft ?? style?.borderLeft,
    paddingTop: directStyle?.paddingTop ?? style?.paddingTop,
    paddingRight: directStyle?.paddingRight ?? style?.paddingRight,
    paddingBottom: directStyle?.paddingBottom ?? style?.paddingBottom,
    paddingLeft: directStyle?.paddingLeft ?? style?.paddingLeft,
    styleId: effectiveStyleId,
    numbering:
      readDocxNumberingReference(pPr) ??
      resolveDocxParagraphStyleProperty(effectiveStyleId, catalog, 'numbering'),
    tabStops:
      readParagraphTabStops(pPr) ??
      resolveDocxParagraphStyleProperty(effectiveStyleId, catalog, 'tabStops'),
    snapToGrid:
      readOnOff(childByLocalName(pPr, 'snapToGrid')) ??
      resolveDocxParagraphStyleProperty(
        effectiveStyleId,
        catalog,
        'snapToGrid',
      ),
    outlineLevel:
      readDocxOutlineLevel(pPr) ??
      resolveDocxOutlineLevel(effectiveStyleId, catalog),
    isTocStyle: isDocxTocStyle(effectiveStyleId, catalog),
    style: contentStyle,
    paragraphMarkStyle,
  };
}

/** 解析并确定 `resolveRunStyle` 对应的引用或配置。 */
export function resolveRunStyle(
  rPr: Element | null | undefined,
  catalog: DocxStyleCatalog,
  theme: OfficeTheme,
) {
  const styleId =
    attr(childByLocalName(rPr, 'rStyle'), 'w:val') ??
    attr(childByLocalName(rPr, 'rStyle'), 'val');
  return mergeTextStyle(
    catalog.defaults.run,
    resolveDocxStyle(styleId, catalog),
    readTextStyle(rPr, theme),
  );
}

export function inlineInheritedStyle(
  style: DocxTextStyle | undefined,
): DocxTextStyle | undefined {
  if (!style) return undefined;
  const {
    backgroundColor,
    borderTop,
    borderRight,
    borderBottom,
    borderLeft,
    paddingTop,
    paddingRight,
    paddingBottom,
    paddingLeft,
    spacingBefore,
    spacingAfter,
    indentLeft,
    indentRight,
    firstLineIndent,
    lineHeight,
    align,
    ...inlineStyle
  } = style;
  return Object.keys(inlineStyle).length ? inlineStyle : undefined;
}

/** 读取 settings.xml 中声明的东亚标点字符间距压缩方式。 */
function readCharacterSpacingControl(
  entries: OfficeEntryMap,
): DocxCharacterSpacingControl | undefined {
  const xml = readXml(entries, 'word/settings.xml');
  if (!xml) return undefined;
  const value = readVal(
    childByLocalName(parseXml(xml).documentElement, 'characterSpacingControl'),
  );
  return value === 'doNotCompress' ||
    value === 'compressPunctuation' ||
    value === 'compressPunctuationAndJapaneseKana'
    ? value
    : undefined;
}

/** 从 WPS 文档网格推导未显式设置行距的正文行高。 */
function readDocumentGridLineHeight(
  bodyNode: Element | null | undefined,
): number | undefined {
  const docGrid = childByLocalName(
    childByLocalName(bodyNode, 'sectPr'),
    'docGrid',
  );
  const gridType = attr(docGrid, 'w:type') ?? attr(docGrid, 'type');
  if (gridType && gridType !== 'lines' && gridType !== 'linesAndChars')
    return undefined;
  return positiveTwipToPx(
    attr(docGrid, 'w:linePitch') ?? attr(docGrid, 'linePitch'),
  );
}

/** 从 WPS 文档网格与默认段落倍数推导正文行高。 */
function readDefaultGridLineHeight(
  bodyNode: Element | null | undefined,
  styles: DocxStyleCatalog,
): number | undefined {
  const linePitch = readDocumentGridLineHeight(bodyNode);
  if (linePitch === undefined) return undefined;
  const defaultStyle = resolveDocxStyle(
    styles.defaults.paragraphStyleId,
    styles,
  );
  const explicitLineMultiplier =
    defaultStyle?.lineHeight !== undefined && defaultStyle.lineHeight <= 4
      ? defaultStyle.lineHeight
      : undefined;
  if (explicitLineMultiplier === undefined && linePitch > 16.1)
    return undefined;
  const lineMultiplier = explicitLineMultiplier ?? 2;
  // WPS 会把普通正文吸附到文档行网格；紧凑双网格减一像素以抵消浏览器行盒取整。
  const gridLineHeight = linePitch * lineMultiplier;
  return linePitch <= 16.1 ? gridLineHeight - 1 : gridLineHeight;
}

/** 创建 DOCX 主解析上下文时使用的选项。 */
export type CreateDocxParseContextOptions = {
  /** DOCX 主文档中的正文根节点。 */
  bodyNode?: Element | null;
  /** 解析上下文可访问的媒体资源索引。 */
  media?: {
    /** 按压缩包路径索引的媒体资源。 */
    byPath: Record<string, OfficeResourceSource>;
    /** 按资源文件名索引的媒体资源。 */
    byName: Record<string, OfficeResourceSource>;
  };
};

/** 为物化和流式解析建立完全相同的样式、主题、关系及媒体上下文。 */
export function createDocxParseContext(
  entries: OfficeEntryMap,
  options: CreateDocxParseContextOptions = {},
): DocxParseContext {
  const packageState = buildDocxPackageState(entries, options.media);
  const theme = readOfficeTheme(readXml(entries, 'word/theme/theme1.xml'));
  const styles = readDocxStyles(entries, theme);
  return {
    packageState,
    documentRels:
      packageState.relationships['word/_rels/document.xml.rels'] ?? {},
    theme,
    defaultLineHeight: readDefaultGridLineHeight(options.bodyNode, styles),
    documentGridLineHeight: readDocumentGridLineHeight(options.bodyNode),
    characterSpacingControl: readCharacterSpacingControl(entries),
    numbering: readDocxNumbering(entries),
    styles,
    images: [],
    bookmarks: {},
    imageIndex: 0,
    chartIndex: 0,
    shapeIndex: 0,
  };
}
