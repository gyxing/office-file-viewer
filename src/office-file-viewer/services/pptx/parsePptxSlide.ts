import { parseOfficeChartXml } from '../../shared/ooxml/charts';
import {
  resolvePackageMediaRef,
  type OfficeRelationship,
} from '../../shared/ooxml/media';
import { getOfficePartRelationshipsPath } from '../../shared/ooxml/relationships';
import { readOfficeTheme, type OfficeTheme } from '../../shared/ooxml/theme';
import { emuToPx } from '../../shared/ooxml/units';
import { parseWpsWebExtensionChartModel } from '../../shared/ooxml/wpsChart';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  descendantByLocalName,
  descendantsByLocalName,
  matchesLocalName,
  parseXml,
  textContent,
} from '../../shared/ooxml/xml';
import {
  resolveOfficePanoseFontWeight,
  resolveOfficeThemeFontFamily,
} from '../fonts/OfficeFontResolver';
import {
  annotatePresentationImageReuse,
  detectPresentationImageSourceKind,
} from '../presentation/imagePreviewPolicy';
import { getPresentationMediaMimeType } from '../presentation/mediaTypes';
import {
  alphaToOpacity,
  alphaToRatio,
  resolveThemeColor,
  toHexColor,
  transformColor,
} from './colors';
import {
  formatPptxTextField,
  type PptxTextFieldContext,
} from './formatPptxTextField';
import { parsePptxComments } from './parsePptxComments';
import {
  hasPptxHyperlinkAction,
  parsePptxHyperlink,
  type PptxSlideTargetMap,
} from './parsePptxHyperlink';
import { parsePptxMediaElement } from './parsePptxMedia';
import { parsePptxTransition } from './parsePptxTransitions';
import { parsePptxSpeakerNotes } from './parseSpeakerNotes';
import type {
  LayoutDefinition,
  MasterDefinition,
  PptxPackageState as PackageState,
  PlaceholderStyle,
  TableCellStyle,
  TableStyleDefinition,
  TableStyleMap,
} from './PptxPackageContext';
import type {
  ChartElement,
  GradientFill,
  ImageCrop,
  ImageElement,
  PresentationImagePreviewMetadata,
  ReflectionStyle,
  ShadowStyle,
  ShapeElement,
  SlideBackground,
  SlideElement,
  SlideModel,
  TableCell,
  TableElement,
  TextElement,
  TextParagraph,
  TextRun,
  TextStyle,
  ThemeModel,
  UnsupportedElement,
} from './types';

/** DrawingML 文本框未声明左右内边距时采用的 0.1 英寸默认值。 */
const DEFAULT_TEXT_BODY_HORIZONTAL_INSET = emuToPx(91440);

/** DrawingML 文本框未声明上下内边距时采用的 0.05 英寸默认值。 */
const DEFAULT_TEXT_BODY_VERTICAL_INSET = emuToPx(45720);

/** DrawingML 所有样式源均未声明字号时采用 PowerPoint 的 18pt 缺省值。 */
const DEFAULT_PRESENTATION_FONT_SIZE = 24;

/** Office 百分比行距基于字体常规行高，CSS 无单位行高则直接基于字号。 */
const OFFICE_FONT_LINE_HEIGHT_RATIO = 1.2;

function emuValue(node: Element | null, name: string) {
  const value = attr(node, name);
  return value ? emuToPx(Number(value)) : undefined;
}

function pointToPx(point?: string) {
  if (!point) return undefined;
  const value = Number(point);
  if (!Number.isFinite(value)) return undefined;
  return (value / 100) * (96 / 72);
}

function pctToLineHeightRatio(value?: string) {
  if (!value) return undefined;
  const next = Number(value);
  if (!Number.isFinite(next)) return undefined;
  return (next / 100000) * OFFICE_FONT_LINE_HEIGHT_RATIO;
}

function clamp01(value: number) {
  return Math.max(0, Math.min(1, value));
}

function boolAttr(node: Element | null, name: string) {
  const value = attr(node, name);
  if (value === undefined) return undefined;
  return value === '1' || value === 'true';
}

/** 合并 `mergeTextStyles` 接收的多份数据。 */
function mergeTextStyles(...styles: Array<TextStyle | undefined>) {
  return styles.reduce<TextStyle>((acc, style) => {
    if (!style) return acc;
    return {
      ...acc,
      ...Object.fromEntries(
        Object.entries(style).filter(([, value]) => value !== undefined),
      ),
    };
  }, {});
}

/** 合并 `mergePlaceholderStyle` 接收的多份数据。 */
function mergePlaceholderStyle(
  base?: PlaceholderStyle,
  override?: PlaceholderStyle,
): PlaceholderStyle {
  if (!base && !override) return {};
  if (!base) return { ...override };
  if (!override) return { ...base };
  // 版式经常只补充文本样式而省略坐标；省略字段不能清空母版中已经定义的几何信息。
  const definedOverride = Object.fromEntries(
    Object.entries(override).filter(([, value]) => value !== undefined),
  ) as Partial<PlaceholderStyle>;
  const levels = { ...(base.levels ?? {}) };
  Object.entries(override.levels ?? {}).forEach(([level, style]) => {
    const levelIndex = Number(level);
    // 空的版式级别样式不能整项覆盖母版，否则会丢失对齐、缩进等继承属性。
    levels[levelIndex] = mergeTextStyles(base.levels?.[levelIndex], style);
  });
  return {
    ...base,
    ...definedOverride,
    fill: override.fill !== undefined ? override.fill : base.fill,
    fillOpacity:
      override.fillOpacity !== undefined
        ? override.fillOpacity
        : base.fillOpacity,
    stroke: override.stroke !== undefined ? override.stroke : base.stroke,
    strokeOpacity:
      override.strokeOpacity !== undefined
        ? override.strokeOpacity
        : base.strokeOpacity,
    strokeWidth:
      override.strokeWidth !== undefined
        ? override.strokeWidth
        : base.strokeWidth,
    strokeDash:
      override.strokeDash !== undefined ? override.strokeDash : base.strokeDash,
    shadow: override.shadow ?? base.shadow,
    reflection: override.reflection ?? base.reflection,
    text: mergeTextStyles(base.text, override.text),
    body: mergeTextStyles(base.body, override.body),
    levels,
  };
}

function readBodyPrStyle(bodyPr: Element | null): TextStyle {
  return {
    verticalAlign:
      attr(bodyPr, 'anchor') === 'ctr'
        ? 'middle'
        : attr(bodyPr, 'anchor') === 'b'
        ? 'bottom'
        : 'top',
    writingMode:
      attr(bodyPr, 'vert') === 'vert'
        ? 'vertical-rl'
        : attr(bodyPr, 'vert') === 'vert270'
        ? 'vertical-lr'
        : 'horizontal-tb',
    fit: childByLocalName(bodyPr, 'spAutoFit')
      ? 'resizeShape'
      : childByLocalName(bodyPr, 'normAutofit')
      ? 'shrinkText'
      : childByLocalName(bodyPr, 'noAutofit')
      ? 'none'
      : undefined,
    marginLeft: attr(bodyPr, 'lIns')
      ? emuToPx(Number(attr(bodyPr, 'lIns')))
      : undefined,
    marginRight: attr(bodyPr, 'rIns')
      ? emuToPx(Number(attr(bodyPr, 'rIns')))
      : undefined,
    marginTop: attr(bodyPr, 'tIns')
      ? emuToPx(Number(attr(bodyPr, 'tIns')))
      : undefined,
    marginBottom: attr(bodyPr, 'bIns')
      ? emuToPx(Number(attr(bodyPr, 'bIns')))
      : undefined,
  };
}

function applyDefaultTextBodyInsets(style: TextStyle): TextStyle {
  return {
    ...style,
    marginLeft: style.marginLeft ?? DEFAULT_TEXT_BODY_HORIZONTAL_INSET,
    marginRight: style.marginRight ?? DEFAULT_TEXT_BODY_HORIZONTAL_INSET,
    marginTop: style.marginTop ?? DEFAULT_TEXT_BODY_VERTICAL_INSET,
    marginBottom: style.marginBottom ?? DEFAULT_TEXT_BODY_VERTICAL_INSET,
  };
}

function readDefaultRunStyle(
  node: Element | null,
  theme: ThemeModel,
): TextStyle {
  if (!node) return {};
  const solidFill = childByLocalName(node, 'solidFill');
  const textFillNode = solidFill ?? childByLocalName(node, 'gradFill');
  const latinFontNode = childByLocalName(node, 'latin');
  const eastAsiaFontNode = childByLocalName(node, 'ea');
  const complexScriptFontNode = childByLocalName(node, 'cs');
  const fontNode = latinFontNode ?? eastAsiaFontNode ?? complexScriptFontNode;
  const themeFonts = {
    majorFont: theme.fontScheme.majorLatin ?? theme.fontScheme.majorFont,
    minorFont: theme.fontScheme.minorLatin ?? theme.fontScheme.minorFont,
    majorEastAsiaFont:
      theme.fontScheme.majorEastAsia ?? theme.fontScheme.majorFont,
    minorEastAsiaFont:
      theme.fontScheme.minorEastAsia ?? theme.fontScheme.minorFont,
  };
  const fontFamily = resolveOfficeThemeFontFamily(
    attr(fontNode, 'typeface'),
    themeFonts,
  );
  const eastAsiaFontFamily = resolveOfficeThemeFontFamily(
    attr(eastAsiaFontNode, 'typeface'),
    themeFonts,
  );
  const capitalization = attr(node, 'cap');
  return {
    fontFamily,
    eastAsiaFontFamily,
    fontSize: pointToPx(attr(node, 'sz')),
    fontWeight: resolveOfficePanoseFontWeight(attr(fontNode, 'panose')),
    bold: boolAttr(node, 'b'),
    italic: boolAttr(node, 'i'),
    underline: attr(node, 'u') === 'sng' || attr(node, 'u') === '1',
    strike:
      attr(node, 'strike') === 'dblStrike'
        ? 'dblStrike'
        : attr(node, 'strike') === 'sngStrike'
        ? 'sngStrike'
        : attr(node, 'strike') === 'none'
        ? 'none'
        : undefined,
    smallCaps:
      capitalization === 'small'
        ? true
        : capitalization === 'none'
        ? false
        : boolAttr(node, 'smCap'),
    allCaps:
      capitalization === 'all'
        ? true
        : capitalization === 'none'
        ? false
        : boolAttr(node, 'cap'),
    color: parseColorNode(solidFill, theme),
    textFill: parsePaintNode(textFillNode, theme),
    opacity: parseAlphaNode(textFillNode),
    reflection: parseReflectionNode(
      childByLocalName(node, 'effectLst') ??
        childByLocalName(node, 'effectDag'),
    ),
    charSpace: pointToPx(attr(node, 'spc')),
    baseline: attr(node, 'baseline')
      ? Number(attr(node, 'baseline')) / 1000
      : undefined,
  };
}

function readParagraphLevelStyle(
  node: Element | null,
  theme: ThemeModel,
): TextStyle {
  if (!node) return {};
  const solidFill = childByLocalName(node, 'solidFill');
  const bulletChar = attr(childByLocalName(node, 'buChar'), 'char');
  const bulletFontFamily = resolveOfficeThemeFontFamily(
    attr(childByLocalName(node, 'buFont'), 'typeface'),
    {
      majorFont: theme.fontScheme.majorLatin ?? theme.fontScheme.majorFont,
      minorFont: theme.fontScheme.minorLatin ?? theme.fontScheme.minorFont,
      majorEastAsiaFont:
        theme.fontScheme.majorEastAsia ?? theme.fontScheme.majorFont,
      minorEastAsiaFont:
        theme.fontScheme.minorEastAsia ?? theme.fontScheme.minorFont,
    },
  );
  const bulletSize = pointToPx(attr(childByLocalName(node, 'buSzPts'), 'val'));
  const bulletColorNode = childByLocalName(node, 'buClr');
  const bulletColor = parseColorNode(
    childByLocalName(bulletColorNode, 'solidFill') ?? bulletColorNode,
    theme,
  );
  const bulletNone = Boolean(childByLocalName(node, 'buNone'));
  const lineSpace = childByLocalName(node, 'lnSpc');
  const spcPct = childByLocalName(lineSpace, 'spcPct');
  const spcPts = childByLocalName(lineSpace, 'spcPts');
  const spcBef = childByLocalName(node, 'spcBef');
  const spcAft = childByLocalName(node, 'spcAft');
  const spaceBefore = pointToPx(
    attr(childByLocalName(spcBef, 'spcPts'), 'val'),
  );
  const spaceAfter = pointToPx(attr(childByLocalName(spcAft, 'spcPts'), 'val'));

  return {
    align:
      attr(node, 'algn') === 'l'
        ? 'left'
        : attr(node, 'algn') === 'ctr'
        ? 'center'
        : attr(node, 'algn') === 'r'
        ? 'right'
        : attr(node, 'algn') === 'just'
        ? 'justify'
        : undefined,
    lineHeight:
      pctToLineHeightRatio(attr(spcPct, 'val')) ??
      pointToPx(attr(spcPts, 'val')),
    marginLeft: attr(node, 'marL')
      ? emuToPx(Number(attr(node, 'marL')))
      : undefined,
    textIndent: attr(node, 'indent')
      ? emuToPx(Number(attr(node, 'indent')))
      : undefined,
    spaceBefore,
    spaceAfter,
    color: parseColorNode(solidFill, theme),
    bullet:
      bulletChar || bulletNone
        ? {
            char: bulletChar,
            fontFamily: bulletFontFamily,
            color: bulletColor,
            size: bulletSize,
            none: bulletNone,
          }
        : undefined,
  };
}

function readLevelStyles(txBody: Element | null, theme: ThemeModel) {
  const listStyle = childByLocalName(txBody, 'lstStyle');
  const levels: Record<number, TextStyle> = {};
  for (let level = 1; level <= 9; level += 1) {
    const node = childByLocalName(listStyle, `lvl${level}pPr`);
    if (!node) continue;
    levels[level - 1] = mergeTextStyles(
      readParagraphLevelStyle(node, theme),
      readDefaultRunStyle(childByLocalName(node, 'defRPr'), theme),
    );
  }
  return levels;
}

function readTextStyleFamily(
  styleNode: Element | null,
  theme: ThemeModel,
): PlaceholderStyle {
  if (!styleNode) return {};
  const defaultParagraph = childByLocalName(styleNode, 'defPPr');
  const defaultRun = childByLocalName(defaultParagraph, 'defRPr');
  const body = mergeTextStyles(
    readParagraphLevelStyle(defaultParagraph, theme),
    readBodyPrStyle(null),
  );
  const text = mergeTextStyles(body, readDefaultRunStyle(defaultRun, theme));
  const levels: Record<number, TextStyle> = {};

  for (let level = 1; level <= 9; level += 1) {
    const node = childByLocalName(styleNode, `lvl${level}pPr`);
    if (!node) continue;
    levels[level - 1] = mergeTextStyles(
      readParagraphLevelStyle(node, theme),
      readDefaultRunStyle(childByLocalName(node, 'defRPr'), theme),
    );
  }

  return { text, body, levels };
}

/** 读取 presentation.xml 中适用于普通文本框的文稿级默认文字样式。 */
export function readPptxDefaultTextStyle(
  presentationXml: string,
  theme: ThemeModel,
): PlaceholderStyle {
  if (!presentationXml) return {};
  const document = parseXml(presentationXml);
  return readTextStyleFamily(
    descendantByLocalName(document.documentElement, 'defaultTextStyle'),
    theme,
  );
}

/** 读取母版或版式的文字预设，供演示结构解析复用。 */
export function readPptxTextPresetMap(
  txStyles: Element | null,
  theme: ThemeModel,
) {
  const presets: Record<string, PlaceholderStyle> = {};
  const titleStyle = childByLocalName(txStyles, 'titleStyle');
  const bodyStyle = childByLocalName(txStyles, 'bodyStyle');
  const otherStyle = childByLocalName(txStyles, 'otherStyle');

  const title = readTextStyleFamily(titleStyle, theme);
  const body = readTextStyleFamily(bodyStyle, theme);
  const other = readTextStyleFamily(otherStyle, theme);

  ['title:0', 'ctrTitle:0', 'title:1', 'ctrTitle:1'].forEach((key) => {
    presets[key] = mergePlaceholderStyle(presets[key], title);
  });
  ['subTitle:0', 'subTitle:1', 'body:0'].forEach((key) => {
    presets[key] = mergePlaceholderStyle(presets[key], body);
  });
  ['dt:10', 'ftr:11', 'sldNum:12', 'other:0'].forEach((key) => {
    presets[key] = mergePlaceholderStyle(presets[key], other);
  });

  return presets;
}

/** 读取表格样式中的单元格外观，供样式表解析复用。 */
export function readPptxTableCellStyle(
  node: Element | null | undefined,
  theme: ThemeModel,
): TableCellStyle {
  if (!node) return {};
  const tcStyle = childByLocalName(node, 'tcStyle') ?? node;
  const tcTxStyle = childByLocalName(node, 'tcTxStyle') ?? node;
  const fillNode = childByLocalName(tcStyle, 'fill');
  const solidFill = childByLocalName(fillNode, 'solidFill');
  const noFill = Boolean(childByLocalName(fillNode, 'noFill'));
  const borderNode = childByLocalName(tcStyle, 'tcBdr');
  const borderLine =
    childByLocalName(borderNode, 'ln') ??
    childByLocalName(borderNode, 'left')?.firstElementChild ??
    childByLocalName(borderNode, 'right')?.firstElementChild ??
    childByLocalName(borderNode, 'top')?.firstElementChild ??
    childByLocalName(borderNode, 'bottom')?.firstElementChild ??
    childByLocalName(borderNode, 'insideH')?.firstElementChild ??
    childByLocalName(borderNode, 'insideV')?.firstElementChild;
  const fill =
    noFill || !solidFill ? undefined : parseColorNode(solidFill, theme);
  const borderFill = childByLocalName(borderLine, 'solidFill');
  const textColorNode =
    childByLocalName(tcTxStyle, 'solidFill') ??
    childByLocalName(tcTxStyle, 'srgbClr') ??
    childByLocalName(tcTxStyle, 'schemeClr') ??
    childByLocalName(tcTxStyle, 'prstClr');
  return {
    text: mergeTextStyles(readDefaultRunStyle(tcTxStyle, theme), {
      color: parseColorNode(textColorNode, theme),
      opacity: parseAlphaNode(textColorNode),
    }),
    backgroundColor: fill,
    backgroundOpacity: parseAlphaNode(solidFill),
    borderColor: childByLocalName(borderLine, 'noFill')
      ? null
      : parseColorNode(borderFill ?? borderLine, theme),
    borderOpacity: parseAlphaNode(borderFill ?? borderLine),
    borderWidth: attr(borderLine, 'w')
      ? Number(attr(borderLine, 'w')) / 12700
      : undefined,
  };
}

/** 合并 `mergeTableCellStyle` 接收的多份数据。 */
function mergeTableCellStyle(...styles: Array<TableCellStyle | undefined>) {
  return styles.reduce<TableCellStyle>((acc, style) => {
    if (!style) return acc;
    return {
      text: mergeTextStyles(acc.text, style.text),
      backgroundColor:
        style.backgroundColor !== undefined
          ? style.backgroundColor
          : acc.backgroundColor,
      backgroundOpacity:
        style.backgroundOpacity !== undefined
          ? style.backgroundOpacity
          : acc.backgroundOpacity,
      borderColor:
        style.borderColor !== undefined ? style.borderColor : acc.borderColor,
      borderOpacity:
        style.borderOpacity !== undefined
          ? style.borderOpacity
          : acc.borderOpacity,
      borderWidth:
        style.borderWidth !== undefined ? style.borderWidth : acc.borderWidth,
    };
  }, {});
}

function readCustomGeometry(spPr: Element | null | undefined) {
  const custGeom = childByLocalName(spPr, 'custGeom');
  if (!custGeom) return {};

  const paths = descendantsByLocalName(custGeom, 'path');
  const pathData: string[] = [];
  let viewBox: string | undefined;

  paths.forEach((pathNode) => {
    const width = attr(pathNode, 'w');
    const height = attr(pathNode, 'h');
    if (!viewBox && width && height) {
      viewBox = `0 0 ${width} ${height}`;
    }

    const commands: string[] = [];
    Array.from(pathNode.children).forEach((child) => {
      if (matchesLocalName(child, 'close')) {
        commands.push('Z');
        return;
      }

      const points = descendantsByLocalName(child, 'pt').map((point) => {
        const x = Number(attr(point, 'x') ?? Number.NaN);
        const y = Number(attr(point, 'y') ?? Number.NaN);
        return Number.isFinite(x) && Number.isFinite(y)
          ? `${x} ${y}`
          : undefined;
      });

      if (matchesLocalName(child, 'moveTo') && points[0]) {
        commands.push(`M ${points[0]}`);
      }
      if (matchesLocalName(child, 'lnTo') && points[0]) {
        commands.push(`L ${points[0]}`);
      }
      if (
        matchesLocalName(child, 'cubicBezTo') &&
        points.length >= 3 &&
        points.every(Boolean)
      ) {
        commands.push(`C ${points.join(' ')}`);
      }
    });

    if (commands.length) {
      pathData.push(commands.join(' '));
    }
  });

  return {
    path: pathData.length ? pathData.join(' ') : undefined,
    viewBox,
  };
}

function readShapeReferenceStyle(
  styleNode: Element | null | undefined,
  theme: ThemeModel,
) {
  const fillRef = childByLocalName(styleNode, 'fillRef');
  const lineRef = childByLocalName(styleNode, 'lnRef');
  const fontRef = childByLocalName(styleNode, 'fontRef');
  const fontColor = parseColorNode(fontRef, theme);
  const useMajorFont = attr(fontRef, 'idx') === 'major';
  const fontTheme = {
    majorFont: theme.fontScheme.majorLatin ?? theme.fontScheme.majorFont,
    minorFont: theme.fontScheme.minorLatin ?? theme.fontScheme.minorFont,
    majorEastAsiaFont:
      theme.fontScheme.majorEastAsia ?? theme.fontScheme.majorFont,
    minorEastAsiaFont:
      theme.fontScheme.minorEastAsia ?? theme.fontScheme.minorFont,
  };

  return {
    fill: parseColorNode(fillRef, theme),
    stroke: parseColorNode(lineRef, theme),
    text: fontRef
      ? {
          color: fontColor,
          textFill: fontColor,
          fontFamily: resolveOfficeThemeFontFamily(
            useMajorFont ? '+mj-lt' : '+mn-lt',
            fontTheme,
          ),
          eastAsiaFontFamily: resolveOfficeThemeFontFamily(
            useMajorFont ? '+mj-ea' : '+mn-ea',
            fontTheme,
          ),
        }
      : undefined,
  };
}

function readShapeVisualStyle(
  spPr: Element | null | undefined,
  theme: ThemeModel,
  styleNode?: Element | null,
) {
  const referenceStyle = readShapeReferenceStyle(styleNode, theme);
  const xfrm = childByLocalName(spPr, 'xfrm');
  const noFill = Boolean(childByLocalName(spPr, 'noFill'));
  const line = childByLocalName(spPr, 'ln');
  const customGeometry = readCustomGeometry(spPr);
  const shape = customGeometry.path
    ? 'path'
    : attr(childByLocalName(spPr, 'prstGeom'), 'prst') ?? 'rect';
  const fillNode =
    childByLocalName(spPr, 'solidFill') ??
    childByLocalName(spPr, 'gradFill') ??
    childByLocalName(spPr, 'pattFill');
  const fill = noFill
    ? null
    : fillNode
    ? parsePaintNode(fillNode, theme)
    : referenceStyle.fill;
  const strokeNone = Boolean(line && childByLocalName(line, 'noFill'));
  const strokeNode =
    childByLocalName(line, 'solidFill') ??
    childByLocalName(line, 'gradFill') ??
    childByLocalName(line, 'pattFill');
  const stroke = strokeNone
    ? null
    : strokeNode
    ? parseColorNode(strokeNode, theme)
    : referenceStyle.stroke;
  const shadow = parseShadowNode(
    childByLocalName(spPr, 'effectLst') ?? childByLocalName(spPr, 'effectDag'),
    theme,
  );
  const reflection = parseReflectionNode(
    childByLocalName(spPr, 'effectLst') ?? childByLocalName(spPr, 'effectDag'),
  );

  return {
    shape,
    fill,
    fillOpacity:
      typeof fill === 'string' ? parseAlphaNode(fillNode) : undefined,
    stroke,
    strokeOpacity: parseAlphaNode(strokeNode ?? line),
    strokeWidth: attr(line, 'w') ? emuToPx(Number(attr(line, 'w'))) : undefined,
    strokeDash: attr(childByLocalName(line, 'prstDash'), 'val') ?? undefined,
    textStyle: referenceStyle.text,
    shadow,
    reflection,
    rotate: attr(xfrm, 'rot') ? Number(attr(xfrm, 'rot')) / 60000 : undefined,
    flipH: attr(xfrm, 'flipH') === '1',
    flipV: attr(xfrm, 'flipV') === '1',
    borderRadius: readBorderRadius(spPr),
    path: customGeometry.path,
    viewBox: customGeometry.viewBox,
  };
}

function readBorderRadius(spPr: Element | null | undefined) {
  const geom = childByLocalName(spPr, 'prstGeom');
  if (attr(geom, 'prst') !== 'roundRect') return undefined;
  const adj = descendantByLocalName(geom, 'gd');
  const value = attr(adj, 'fmla');
  if (!value) return undefined;
  const match = value.match(/val\s+(\d+)/i);
  const ratio = match ? Number(match[1]) / 100000 : undefined;
  return ratio;
}

function colorWithOpacity(color: string, opacity?: number) {
  if (opacity === undefined || opacity >= 1) return color;
  const normalized = color.replace('#', '');
  if (!/^[0-9a-f]{6}$/i.test(normalized)) return color;
  const value = Number.parseInt(normalized, 16);
  const r = (value >> 16) & 255;
  const g = (value >> 8) & 255;
  const b = value & 255;
  return `rgba(${r}, ${g}, ${b}, ${opacity})`;
}

function readGradientFill(
  node: Element | null | undefined,
  theme: ThemeModel,
): GradientFill | undefined {
  if (!node || !matchesLocalName(node, 'gradFill')) return undefined;

  const stops = childrenByLocalName(childByLocalName(node, 'gsLst'), 'gs')
    .map((stop) => {
      const colorNode =
        childByLocalName(stop, 'srgbClr') ??
        childByLocalName(stop, 'schemeClr') ??
        childByLocalName(stop, 'sysClr') ??
        childByLocalName(stop, 'prstClr');
      const color = parseColorNode(colorNode, theme);
      if (!color) return undefined;
      return {
        offset: clamp01(Number(attr(stop, 'pos') ?? 0) / 100000),
        color: colorWithOpacity(color, parseAlphaNode(colorNode)),
      };
    })
    .filter(
      (
        stop,
      ): stop is {
        /** 在所属数据范围中的偏移位置。 */
        offset: number;
        /** 渐变色标对应的 CSS 颜色值。 */
        color: string;
      } => Boolean(stop),
    )
    .sort((a, b) => a.offset - b.offset);

  if (!stops.length) return undefined;

  return {
    type: 'linear',
    angle: Number(attr(childByLocalName(node, 'lin'), 'ang') ?? 0) / 60000,
    stops,
  };
}

function pickGradientColorNode(node: Element | null | undefined) {
  if (!node || !matchesLocalName(node, 'gradFill')) return node;
  const stops = descendantsByLocalName(node, 'gs');
  for (let index = stops.length - 1; index >= 0; index -= 1) {
    const colorNode =
      childByLocalName(stops[index], 'srgbClr') ??
      childByLocalName(stops[index], 'schemeClr') ??
      childByLocalName(stops[index], 'sysClr') ??
      childByLocalName(stops[index], 'prstClr');
    if (colorNode) return colorNode;
  }
  return node;
}

function parsePaintNode(node: Element | null | undefined, theme: ThemeModel) {
  return readGradientFill(node, theme) ?? parseColorNode(node, theme);
}

function parseColorNode(node: Element | null | undefined, theme: ThemeModel) {
  const sourceNode = pickGradientColorNode(node);
  if (!sourceNode) return undefined;
  const srgb = matchesLocalName(sourceNode, 'srgbClr')
    ? sourceNode
    : childByLocalName(sourceNode, 'srgbClr');
  const scheme = matchesLocalName(sourceNode, 'schemeClr')
    ? sourceNode
    : childByLocalName(sourceNode, 'schemeClr');
  const sys = matchesLocalName(sourceNode, 'sysClr')
    ? sourceNode
    : childByLocalName(sourceNode, 'sysClr');
  const prst = matchesLocalName(sourceNode, 'prstClr')
    ? sourceNode
    : childByLocalName(sourceNode, 'prstClr');
  const colorNode = srgb ?? scheme ?? sys ?? prst;
  const base =
    attr(srgb, 'val') ??
    resolveThemeColor(attr(scheme, 'val'), theme) ??
    attr(sys, 'lastClr') ??
    attr(prst, 'val');
  const transforms = descendantsByLocalName(colorNode, 'tint')
    .concat(descendantsByLocalName(colorNode, 'shade'))
    .concat(descendantsByLocalName(colorNode, 'lumMod'))
    .concat(descendantsByLocalName(colorNode, 'lumOff'))
    .map((item) => ({
      type: item.localName,
      val: Number(attr(item, 'val') ?? 0),
    }));
  const raw = transformColor(toHexColor(base), transforms);
  return raw;
}

function parseAlphaNode(node: Element | null | undefined) {
  return alphaToOpacity(
    attr(descendantByLocalName(pickGradientColorNode(node), 'alpha'), 'val'),
  );
}

function parseRatioNode(node: Element | null | undefined) {
  return alphaToRatio(attr(descendantByLocalName(node, 'alpha'), 'val'));
}

function parseShadowNode(
  node: Element | null,
  theme: ThemeModel,
): ShadowStyle | undefined {
  if (!node) return undefined;
  // 阴影参数位于 effectLst/effectDag 内的具体阴影节点，不能从效果容器读取。
  const shadowNode =
    (matchesLocalName(node, 'outerShdw') || matchesLocalName(node, 'innerShdw')
      ? node
      : childByLocalName(node, 'outerShdw') ??
        childByLocalName(node, 'innerShdw') ??
        descendantByLocalName(node, 'outerShdw') ??
        descendantByLocalName(node, 'innerShdw')) ?? undefined;
  if (!shadowNode) return undefined;
  const colorNode =
    childByLocalName(shadowNode, 'srgbClr') ??
    childByLocalName(shadowNode, 'schemeClr');
  const color = parseColorNode(colorNode, theme);
  const opacity = parseRatioNode(colorNode);
  const blur = attr(shadowNode, 'blurRad')
    ? Number(attr(shadowNode, 'blurRad')) / 12700
    : undefined;
  const dist = attr(shadowNode, 'dist')
    ? Number(attr(shadowNode, 'dist')) / 12700
    : undefined;
  const dir = attr(shadowNode, 'dir')
    ? (Number(attr(shadowNode, 'dir')) / 60000) * (Math.PI / 180)
    : undefined;
  return {
    color,
    opacity,
    blur,
    offsetX: dist && dir !== undefined ? Math.cos(dir) * dist : undefined,
    offsetY: dist && dir !== undefined ? Math.sin(dir) * dist : undefined,
  };
}

/** 读取 DrawingML 倒影参数，渲染层再转换为浏览器盒倒影。 */
function parseReflectionNode(
  node: Element | null | undefined,
): ReflectionStyle | undefined {
  if (!node) return undefined;
  const reflectionNode =
    (matchesLocalName(node, 'reflection')
      ? node
      : childByLocalName(node, 'reflection') ??
        descendantByLocalName(node, 'reflection')) ?? undefined;
  if (!reflectionNode) return undefined;
  const ratio = (name: string) => {
    const value = attr(reflectionNode, name);
    if (value === undefined) return undefined;
    const parsed = Number(value) / 100000;
    return Number.isFinite(parsed) ? clamp01(parsed) : undefined;
  };
  const emu = (name: string) => {
    const value = attr(reflectionNode, name);
    if (value === undefined) return undefined;
    const parsed = Number(value) / 12700;
    return Number.isFinite(parsed) ? parsed : undefined;
  };
  const direction = attr(reflectionNode, 'dir');
  return {
    blur: emu('blurRad'),
    distance: emu('dist'),
    direction:
      direction !== undefined && Number.isFinite(Number(direction))
        ? Number(direction) / 60000
        : undefined,
    startOpacity: ratio('stA'),
    endOpacity: ratio('endA'),
    startPosition: ratio('stPos'),
    endPosition: ratio('endPos'),
  };
}

/** 解析并确定 `resolveMediaRef` 对应的引用或配置。 */
function resolveMediaRef(
  target: string | undefined,
  packageState: PackageState,
) {
  return resolvePackageMediaRef(
    target,
    packageState.mediaByPath,
    packageState.mediaByName,
    'ppt',
  );
}

/** 解析并确定 `resolveXmlTarget` 对应的引用或配置。 */
function resolveXmlTarget(
  target: string | undefined,
  packageState: PackageState,
) {
  if (!target) return undefined;
  const normalized = target.replace(/^\.\.\//, '');
  return packageState.entries.get(normalized) ? normalized : target;
}

/** 解析并确定 `resolveWebExtensionSnapshot` 对应的引用或配置。 */
function resolveWebExtensionSnapshot(
  doc: XMLDocument,
  webExtensionPath: string,
  packageState: PackageState,
) {
  const snapshot = descendantByLocalName(doc.documentElement, 'snapshot');
  const embed = attr(snapshot, 'r:embed') ?? attr(snapshot, 'embed');
  const relsPath = webExtensionPath
    .replace(/^ppt\/webExtensions\//, 'ppt/webExtensions/_rels/')
    .concat('.rels');
  const target = embed
    ? packageState.relationships[relsPath]?.[embed]?.target
    : undefined;
  return resolveMediaRef(target, packageState);
}

/** 解析幻灯片、母版或版式的背景。 */
export function readPptxSlideBackground(
  bgNode: Element | null,
  theme: ThemeModel,
  packageState: PackageState,
  slideRels?: Record<string, OfficeRelationship>,
): SlideBackground | undefined {
  if (!bgNode) return undefined;
  const bgPr = childByLocalName(bgNode, 'bgPr');
  if (!bgPr) return undefined;
  const solidFill = childByLocalName(bgPr, 'solidFill');
  const fill = parseColorNode(solidFill, theme);
  const blip = descendantByLocalName(bgPr, 'blip');
  const embed = attr(blip, 'r:embed') ?? attr(blip, 'embed');
  const target = embed ? slideRels?.[embed]?.target : undefined;
  return {
    fill,
    fillOpacity: parseAlphaNode(solidFill),
    imageRef: resolveMediaRef(target, packageState),
  };
}

/** 读取占位符的布局、外观和文本继承信息。 */
export function readPptxPlaceholder(
  node: Element | null,
  theme: ThemeModel,
): PlaceholderStyle {
  if (!node) return {};
  const ph = descendantByLocalName(node, 'ph');
  const spPr = childByLocalName(node, 'spPr');
  const xfrm = childByLocalName(spPr, 'xfrm');
  const off = childByLocalName(xfrm, 'off');
  const ext = childByLocalName(xfrm, 'ext');
  const visual = readShapeVisualStyle(
    spPr,
    theme,
    childByLocalName(node, 'style'),
  );
  const textBody = childByLocalName(node, 'txBody');
  const bodyPr = childByLocalName(textBody, 'bodyPr');
  const defRPr =
    descendantByLocalName(textBody, 'defRPr') ??
    descendantByLocalName(textBody, 'endParaRPr');

  return {
    type: attr(ph, 'type') ?? undefined,
    idx: attr(ph, 'idx') ?? undefined,
    x: emuValue(off, 'x'),
    y: emuValue(off, 'y'),
    width: emuValue(ext, 'cx'),
    height: emuValue(ext, 'cy'),
    ...visual,
    text: mergeTextStyles(
      readDefaultRunStyle(defRPr, theme),
      readBodyPrStyle(bodyPr),
    ),
    body: mergeTextStyles(readBodyPrStyle(bodyPr)),
    levels: readLevelStyles(textBody, theme),
  };
}

/** 解析并确定 `resolvePlaceholderStyle` 对应的引用或配置。 */
function resolvePlaceholderStyle(
  ph: Element,
  placeholderStyles?: Record<string, PlaceholderStyle>,
) {
  if (!placeholderStyles) return undefined;
  const type = attr(ph, 'type') ?? 'body';
  const idx = attr(ph, 'idx') ?? '0';
  const aliases = [
    `${type}:${idx}`,
    `${type}:0`,
    type === 'ctrTitle' ? 'title:0' : undefined,
    type === 'title' ? 'ctrTitle:0' : undefined,
    type === 'subTitle' ? 'body:0' : undefined,
    type === 'subTitle' ? 'title:0' : undefined,
    type === 'dt' ? 'other:0' : undefined,
    type === 'ftr' ? 'other:0' : undefined,
    type === 'sldNum' ? 'other:0' : undefined,
    `body:${idx}`,
    'body:0',
  ].filter((key): key is string => Boolean(key));

  for (const key of aliases) {
    if (placeholderStyles[key]) return placeholderStyles[key];
  }

  return undefined;
}

function transformGroupedElement(
  element: {
    /** 水平方向的坐标或缩放参数。 */
    x: number;
    /** 垂直方向的坐标或缩放参数。 */
    y: number;
    /** 对象宽度，单位为标准化渲染像素。 */
    width: number;
    /** 对象高度，单位为标准化渲染像素。 */
    height: number;
  },
  group: {
    /** 水平方向的坐标或缩放参数。 */
    x: number;
    /** 垂直方向的坐标或缩放参数。 */
    y: number;
    /** 对象宽度，单位为标准化渲染像素。 */
    width: number;
    /** 对象高度，单位为标准化渲染像素。 */
    height: number;
    /** 组合图形子坐标系的水平原点。 */
    childX: number;
    /** 组合图形子坐标系的垂直原点。 */
    childY: number;
    /** 组合图形子坐标系的宽度。 */
    childWidth: number;
    /** 组合图形子坐标系的高度。 */
    childHeight: number;
  },
) {
  const scaleX = group.childWidth ? group.width / group.childWidth : 1;
  const scaleY = group.childHeight ? group.height / group.childHeight : 1;
  return {
    x: group.x + (element.x - group.childX) * scaleX,
    y: group.y + (element.y - group.childY) * scaleY,
    width: element.width * scaleX,
    height: element.height * scaleY,
  };
}

/** 读取图形、图片或图表自身非可视属性中的点击链接。 */
function parsePptxObjectHyperlink(
  node: Element,
  relationships: Record<string, OfficeRelationship>,
  slideTargets?: PptxSlideTargetMap,
) {
  return parsePptxHyperlink(
    descendantByLocalName(node, 'cNvPr'),
    relationships,
    slideTargets,
  );
}

/** 读取 PresentationML 非可视属性中的源对象编号，供现代批注精确定位。 */
function readPptxSourceObjectId(node: Element) {
  return attr(descendantByLocalName(node, 'cNvPr'), 'id') ?? undefined;
}

function parseGroupElement(
  node: Element,
  index: number,
  theme: ThemeModel,
  packageState: PackageState,
  rels: Record<string, OfficeRelationship>,
  sourcePrefix: string,
  placeholderStyles?: Record<string, PlaceholderStyle>,
  tableStyles?: TableStyleMap,
  includePlaceholders = true,
  fieldContext?: PptxTextFieldContext,
  slideTargets?: PptxSlideTargetMap,
  defaultTextStyle?: PlaceholderStyle,
) {
  const spPr = childByLocalName(node, 'grpSpPr');
  const xfrm = childByLocalName(spPr, 'xfrm');
  const offsetX = emuValue(childByLocalName(xfrm, 'off'), 'x') ?? 0;
  const offsetY = emuValue(childByLocalName(xfrm, 'off'), 'y') ?? 0;
  const childX = emuValue(childByLocalName(xfrm, 'chOff'), 'x') ?? 0;
  const childY = emuValue(childByLocalName(xfrm, 'chOff'), 'y') ?? 0;
  const childWidth =
    emuValue(childByLocalName(xfrm, 'chExt'), 'cx') ??
    emuValue(childByLocalName(xfrm, 'ext'), 'cx') ??
    0;
  const childHeight =
    emuValue(childByLocalName(xfrm, 'chExt'), 'cy') ??
    emuValue(childByLocalName(xfrm, 'ext'), 'cy') ??
    0;
  const width = emuValue(childByLocalName(xfrm, 'ext'), 'cx') ?? 0;
  const height = emuValue(childByLocalName(xfrm, 'ext'), 'cy') ?? 0;
  const inner = childByLocalName(node, 'spTree') ?? node;
  const childElements = parsePptxVisualTree(
    inner,
    theme,
    packageState,
    rels,
    `${sourcePrefix}-group-${index}`,
    placeholderStyles,
    tableStyles,
    includePlaceholders,
    fieldContext,
    slideTargets,
    defaultTextStyle,
  );
  return childElements.map((element) => {
    const translated = transformGroupedElement(element, {
      x: offsetX,
      y: offsetY,
      width,
      height,
      childX,
      childY,
      childWidth,
      childHeight,
    });
    return {
      ...element,
      id: `${sourcePrefix}-group-${index}-${element.id}`,
      ...translated,
    };
  });
}

/** 解析 PPTX 视觉树，并保持包内绘制顺序。 */
export function parsePptxVisualTree(
  spTree: Element | null,
  theme: ThemeModel,
  packageState: PackageState,
  rels: Record<string, OfficeRelationship>,
  sourcePrefix: string,
  placeholderStyles?: Record<string, PlaceholderStyle>,
  tableStyles?: TableStyleMap,
  includePlaceholders = true,
  fieldContext?: PptxTextFieldContext,
  slideTargets?: PptxSlideTargetMap,
  defaultTextStyle?: PlaceholderStyle,
) {
  const elements: SlideElement[] = [];
  const nodes = childrenByLocalName(spTree, 'sp')
    .concat(childrenByLocalName(spTree, 'pic'))
    .concat(childrenByLocalName(spTree, 'graphicFrame'))
    .concat(childrenByLocalName(spTree, 'cxnSp'))
    .concat(childrenByLocalName(spTree, 'grpSp'))
    .sort(
      (a, b) =>
        Array.from(spTree?.children ?? []).indexOf(a) -
        Array.from(spTree?.children ?? []).indexOf(b),
    );

  nodes.forEach((node, elementIndex) => {
    const sourceObjectId = readPptxSourceObjectId(node);
    if (matchesLocalName(node, 'pic')) {
      const media = parsePptxMediaElement(
        node,
        elementIndex,
        packageState,
        rels,
      );
      if (media) {
        media.id = `${sourcePrefix}-${media.id}`;
        media.sourceObjectId = sourceObjectId;
        media.hyperlink = parsePptxObjectHyperlink(node, rels, slideTargets);
        elements.push(media);
        return;
      }
      const chart = parseWpsWebExtensionChart(
        node,
        elementIndex,
        packageState,
        rels,
      );
      if (chart) {
        chart.hyperlink = parsePptxObjectHyperlink(node, rels, slideTargets);
        chart.id = `${sourcePrefix}-${chart.id}`;
        chart.sourceObjectId = sourceObjectId;
        elements.push(chart);
        return;
      }
      const image = parseImageElement(node, elementIndex, packageState, rels);
      image.hyperlink = parsePptxObjectHyperlink(node, rels, slideTargets);
      image.id = `${sourcePrefix}-${image.id}`;
      image.sourceObjectId = sourceObjectId;
      elements.push(image);
      return;
    }

    if (matchesLocalName(node, 'graphicFrame')) {
      const chart = parseChartElement(
        node,
        elementIndex,
        theme,
        packageState,
        rels,
      );
      const tbl = descendantByLocalName(node, 'tbl');
      const element =
        chart ??
        (tbl
          ? parseTableElement(
              node,
              elementIndex,
              theme,
              tableStyles,
              rels,
              slideTargets,
            )
          : parseUnsupportedElement(elementIndex, 'Unsupported graphic frame'));
      element.hyperlink = parsePptxObjectHyperlink(node, rels, slideTargets);
      element.id = `${sourcePrefix}-${element.id}`;
      element.sourceObjectId = sourceObjectId;
      elements.push(element);
      return;
    }

    if (matchesLocalName(node, 'grpSp')) {
      const groupElements = parseGroupElement(
        node,
        elementIndex,
        theme,
        packageState,
        rels,
        sourcePrefix,
        placeholderStyles,
        tableStyles,
        includePlaceholders,
        fieldContext,
        slideTargets,
        defaultTextStyle,
      );
      elements.push(...groupElements);
      return;
    }

    const ph = descendantByLocalName(node, 'ph');
    if (ph && !includePlaceholders) {
      return;
    }
    const shapeStyleNode = childByLocalName(node, 'style');
    const inherited = ph
      ? resolvePlaceholderStyle(ph, placeholderStyles)
      : shapeStyleNode
      ? mergePlaceholderStyle(defaultTextStyle, placeholderStyles?.['other:0'])
      : defaultTextStyle;
    const txBody = childByLocalName(node, 'txBody');
    const hasText = Boolean(txBody);
    const hasTextContent = Boolean(
      txBody &&
        descendantsByLocalName(txBody, 't').some((textNode) =>
          Boolean(textNode.textContent?.trim()),
        ),
    );
    const visualNode = childByLocalName(node, 'spPr');
    const imageFill = descendantByLocalName(visualNode, 'blipFill');
    const visual = visualNode
      ? readShapeVisualStyle(visualNode, theme, shapeStyleNode)
      : undefined;
    const hasVisibleVisual = Boolean(
      imageFill ||
        (visual &&
          ((visual.fill !== undefined && visual.fill !== null) ||
            (visual.stroke !== undefined && visual.stroke !== null) ||
            visual.shadow)),
    );
    if (ph && !placeholderStyles && !hasText && !hasVisibleVisual) {
      return;
    }
    if (
      ph &&
      !hasText &&
      !inherited?.fill &&
      !inherited?.stroke &&
      !inherited?.shadow
    ) {
      return;
    }
    if (imageFill) {
      const image = parseImageElement(
        node,
        elementIndex,
        packageState,
        rels,
        false,
      );
      image.hyperlink = parsePptxObjectHyperlink(node, rels, slideTargets);
      image.id = `${sourcePrefix}-image-fill-${elementIndex}`;
      image.sourceObjectId = sourceObjectId;
      elements.push(image);
      // 图片填充是形状的视觉底层；仅当文本框确有内容时继续叠加文字。
      if (!hasTextContent) return;
    }

    elements.push(
      hasText
        ? parsePptxTextElement(
            node,
            elementIndex,
            theme,
            inherited,
            fieldContext,
            rels,
            slideTargets,
          )
        : parseShapeElement(
            node,
            elementIndex,
            theme,
            inherited,
            rels,
            slideTargets,
          ),
    );
    elements[elements.length - 1].id = `${sourcePrefix}-${
      elements[elements.length - 1].id
    }`;
    elements[elements.length - 1].sourceObjectId = sourceObjectId;
  });

  return elements;
}

function resolvePlaceholderPreset(
  key: string,
  presets: Record<string, PlaceholderStyle>,
) {
  const type = key.split(':', 1)[0];
  const presetKey =
    type === 'title' || type === 'ctrTitle'
      ? 'title:0'
      : type === 'dt' || type === 'ftr' || type === 'sldNum'
      ? 'other:0'
      : type === 'other'
      ? 'other:0'
      : 'body:0';
  return mergePlaceholderStyle(presets[presetKey], presets[key]);
}

/** 页脚类占位符在母版和版式中可能使用不同 idx，需要按类型继续继承几何信息。 */
function resolveFooterPlaceholder(
  key: string,
  placeholders: Record<string, PlaceholderStyle>,
) {
  if (placeholders[key]) return placeholders[key];
  const type = key.split(':', 1)[0];
  if (type !== 'dt' && type !== 'ftr' && type !== 'sldNum') return undefined;
  const matchingKey = Object.keys(placeholders).find(
    (candidate) => candidate.split(':', 1)[0] === type,
  );
  return matchingKey ? placeholders[matchingKey] : undefined;
}

/** 按母版预设、母版占位符、版式预设、版式占位符的 Office 继承顺序组装样式。 */
function buildPlaceholderStyles(
  master?: MasterDefinition,
  layout?: LayoutDefinition,
  defaultTextStyle?: PlaceholderStyle,
) {
  const styles: Record<string, PlaceholderStyle> = {};
  const keys = new Set([
    ...Object.keys(master?.textPresets ?? {}),
    ...Object.keys(master?.placeholders ?? {}),
    ...Object.keys(layout?.textPresets ?? {}),
    ...Object.keys(layout?.placeholders ?? {}),
  ]);
  keys.forEach((key) => {
    const masterStyle = mergePlaceholderStyle(
      resolvePlaceholderPreset(key, master?.textPresets ?? {}),
      resolveFooterPlaceholder(key, master?.placeholders ?? {}),
    );
    const layoutStyle = mergePlaceholderStyle(
      resolvePlaceholderPreset(key, layout?.textPresets ?? {}),
      resolveFooterPlaceholder(key, layout?.placeholders ?? {}),
    );
    styles[key] = mergePlaceholderStyle(
      mergePlaceholderStyle(defaultTextStyle, masterStyle),
      layoutStyle,
    );
  });
  return styles;
}

/** 将单个 DrawingML 文本形状转换为公共演示文本模型。 */
export function parsePptxTextElement(
  node: Element,
  index: number,
  theme: ThemeModel,
  inherited?: PlaceholderStyle,
  fieldContext?: PptxTextFieldContext,
  relationships: Record<string, OfficeRelationship> = {},
  slideTargets?: PptxSlideTargetMap,
): TextElement {
  const spPr = childByLocalName(node, 'spPr');
  const xfrm = childByLocalName(spPr, 'xfrm');
  const off = childByLocalName(xfrm, 'off');
  const ext = childByLocalName(xfrm, 'ext');
  const txBody = childByLocalName(node, 'txBody');
  const bodyPr = childByLocalName(txBody, 'bodyPr');
  const visual = readShapeVisualStyle(
    spPr,
    theme,
    childByLocalName(node, 'style'),
  );
  const localLevels = readLevelStyles(txBody, theme);
  const bodyStyle = applyDefaultTextBodyInsets(
    mergeTextStyles(inherited?.body, readBodyPrStyle(bodyPr)),
  );

  const paragraphs: TextParagraph[] = childrenByLocalName(txBody, 'p').map(
    (paragraphNode) => {
      const paragraphProps = childByLocalName(paragraphNode, 'pPr');
      const level = Number(attr(paragraphProps, 'lvl') ?? 0);
      // 文本样式来源很多：母版占位符、layout、段落级别、run 自身，需要按 Office 优先级合并。
      const levelStyle = mergeTextStyles(
        inherited?.levels?.[level],
        localLevels[level],
        readParagraphLevelStyle(paragraphProps, theme),
      );
      const defaultRunStyle = mergeTextStyles(
        inherited?.text,
        inherited?.body,
        inherited?.levels?.[level],
        localLevels[level],
        visual.textStyle,
        readDefaultRunStyle(childByLocalName(paragraphProps, 'defRPr'), theme),
      );
      const endParagraphStyle = mergeTextStyles(
        defaultRunStyle,
        readDefaultRunStyle(
          childByLocalName(paragraphNode, 'endParaRPr'),
          theme,
        ),
      );

      const runs: TextRun[] = [];
      Array.from(paragraphNode.children).forEach((child) => {
        if (child.localName === 'r') {
          const runProps = childByLocalName(child, 'rPr');
          runs.push({
            text: textContent(childByLocalName(child, 't')),
            hyperlink: parsePptxHyperlink(
              runProps,
              relationships,
              slideTargets,
            ),
            style: mergeTextStyles(
              defaultRunStyle,
              readDefaultRunStyle(runProps, theme),
            ),
          });
          return;
        }

        if (child.localName === 'fld') {
          const runProps = childByLocalName(child, 'rPr');
          const storedText =
            textContent(childByLocalName(child, 't')) ||
            child.textContent ||
            '';
          runs.push({
            text: formatPptxTextField(child, storedText, fieldContext),
            fieldType: attr(child, 'type'),
            hyperlink: parsePptxHyperlink(
              runProps,
              relationships,
              slideTargets,
            ),
            style: mergeTextStyles(
              defaultRunStyle,
              readDefaultRunStyle(runProps, theme),
            ),
          });
          return;
        }

        if (child.localName === 'br') {
          const runProps = childByLocalName(child, 'rPr');
          runs.push({
            text: '\n',
            style: mergeTextStyles(
              defaultRunStyle,
              readDefaultRunStyle(runProps, theme),
            ),
          });
        }
      });

      if (!runs.length) {
        runs.push({
          text: paragraphNode.textContent ?? '',
          style: endParagraphStyle,
        });
      }

      return {
        level,
        runs,
        style: mergeTextStyles(levelStyle, {
          align: levelStyle.align ?? bodyStyle.align,
        }),
        bullet: levelStyle.bullet,
      };
    },
  );

  const firstRunStyle =
    paragraphs.flatMap((paragraph) => paragraph.runs).find(Boolean)?.style ??
    {};
  const fallbackStyle = mergeTextStyles(
    inherited?.text,
    bodyStyle,
    visual.textStyle,
  );

  const objectProperties = descendantByLocalName(node, 'cNvPr');
  const hyperlink = parsePptxHyperlink(
    objectProperties,
    relationships,
    slideTargets,
  );
  return {
    id: `text-${index}`,
    type: 'text',
    x: emuValue(off, 'x') ?? inherited?.x ?? 0,
    y: emuValue(off, 'y') ?? inherited?.y ?? 0,
    width: emuValue(ext, 'cx') ?? inherited?.width ?? 0,
    height: emuValue(ext, 'cy') ?? inherited?.height ?? 0,
    rotate: visual.rotate,
    flipH: visual.flipH,
    flipV: visual.flipV,
    placeholderType: inherited?.type,
    placeholderIdx: inherited?.idx,
    hyperlink,
    hyperlinkSourceType: hyperlink
      ? hasPptxHyperlinkAction(objectProperties)
        ? 'button'
        : 'shape'
      : undefined,
    paragraphs,
    shape: visual.shape,
    path: visual.path,
    viewBox: visual.viewBox,
    fill: visual.fill !== undefined ? visual.fill : inherited?.fill,
    fillOpacity: visual.fillOpacity ?? inherited?.fillOpacity,
    stroke: visual.stroke !== undefined ? visual.stroke : inherited?.stroke,
    strokeOpacity: visual.strokeOpacity ?? inherited?.strokeOpacity,
    strokeWidth: visual.strokeWidth ?? inherited?.strokeWidth,
    strokeDash: visual.strokeDash ?? inherited?.strokeDash,
    shadow: visual.shadow ?? inherited?.shadow,
    reflection: visual.reflection ?? inherited?.reflection,
    borderRadius: visual.borderRadius,
    boxStyle: {
      fontFamily: firstRunStyle.fontFamily ?? fallbackStyle.fontFamily,
      eastAsiaFontFamily:
        firstRunStyle.eastAsiaFontFamily ?? fallbackStyle.eastAsiaFontFamily,
      fontSize:
        firstRunStyle.fontSize ??
        fallbackStyle.fontSize ??
        DEFAULT_PRESENTATION_FONT_SIZE,
      // 行内格式已经落在各 run 上，文本框只保留继承默认值，避免首个粗体 run 污染后续普通正文。
      fontWeight: fallbackStyle.fontWeight,
      bold: fallbackStyle.bold,
      italic: fallbackStyle.italic,
      underline: fallbackStyle.underline,
      color: fallbackStyle.color,
      opacity: fallbackStyle.opacity,
      align:
        paragraphs[0]?.style?.align ?? bodyStyle.align ?? fallbackStyle.align,
      lineHeight: fallbackStyle.lineHeight,
      marginLeft: bodyStyle.marginLeft,
      marginRight: bodyStyle.marginRight,
      marginTop: bodyStyle.marginTop,
      marginBottom: bodyStyle.marginBottom,
      verticalAlign:
        bodyStyle.verticalAlign ?? fallbackStyle.verticalAlign ?? 'top',
      writingMode: bodyStyle.writingMode ?? fallbackStyle.writingMode,
      fit: bodyStyle.fit ?? fallbackStyle.fit,
    },
  };
}

function parseShapeElement(
  node: Element,
  index: number,
  theme: ThemeModel,
  inherited?: PlaceholderStyle,
  relationships: Record<string, OfficeRelationship> = {},
  slideTargets?: PptxSlideTargetMap,
): ShapeElement {
  const spPr = childByLocalName(node, 'spPr');
  const xfrm = childByLocalName(spPr, 'xfrm');
  const ph = descendantByLocalName(node, 'ph');
  const visual = readShapeVisualStyle(
    spPr,
    theme,
    childByLocalName(node, 'style'),
  );
  const objectProperties = descendantByLocalName(node, 'cNvPr');
  const hyperlink = parsePptxHyperlink(
    objectProperties,
    relationships,
    slideTargets,
  );

  return {
    id: `shape-${index}`,
    type: 'shape',
    shape: visual.shape,
    path: visual.path,
    viewBox: visual.viewBox,
    x: emuValue(childByLocalName(xfrm, 'off'), 'x') ?? inherited?.x ?? 0,
    y: emuValue(childByLocalName(xfrm, 'off'), 'y') ?? inherited?.y ?? 0,
    width:
      emuValue(childByLocalName(xfrm, 'ext'), 'cx') ?? inherited?.width ?? 0,
    height:
      emuValue(childByLocalName(xfrm, 'ext'), 'cy') ?? inherited?.height ?? 0,
    rotate: visual.rotate,
    flipH: visual.flipH,
    flipV: visual.flipV,
    fill: visual.fill !== undefined ? visual.fill : inherited?.fill,
    fillOpacity: visual.fillOpacity ?? inherited?.fillOpacity,
    stroke: visual.stroke !== undefined ? visual.stroke : inherited?.stroke,
    strokeOpacity: visual.strokeOpacity ?? inherited?.strokeOpacity,
    strokeWidth: visual.strokeWidth ?? inherited?.strokeWidth,
    opacity: visual.fillOpacity,
    strokeDash: visual.strokeDash ?? inherited?.strokeDash,
    shadow: visual.shadow ?? inherited?.shadow,
    reflection: visual.reflection ?? inherited?.reflection,
    placeholderType: attr(ph, 'type') ?? inherited?.type,
    placeholderIdx: attr(ph, 'idx') ?? inherited?.idx,
    hyperlink,
    hyperlinkSourceType: hyperlink
      ? hasPptxHyperlinkAction(objectProperties)
        ? 'button'
        : 'shape'
      : undefined,
    borderRadius: visual.borderRadius,
  };
}

function parseImageElement(
  node: Element,
  index: number,
  packageState: PackageState,
  slideRels: Record<string, OfficeRelationship>,
  previewable = true,
): ImageElement {
  const xfrm = childByLocalName(childByLocalName(node, 'spPr') ?? node, 'xfrm');
  const blip = descendantByLocalName(node, 'blip');
  const svgBlip = descendantByLocalName(node, 'svgBlip');
  const blipFill = descendantByLocalName(node, 'blipFill');
  const srcRect = childByLocalName(blipFill, 'srcRect');
  const embed =
    attr(svgBlip, 'r:embed') ??
    attr(svgBlip, 'embed') ??
    attr(blip, 'r:embed') ??
    attr(blip, 'embed');
  const target = embed ? slideRels[embed]?.target : undefined;
  const resolved = resolveMediaRef(target, packageState);
  const objectProperties = descendantByLocalName(node, 'cNvPr');
  const sourceEntry = target ? packageState.entries.get(target) : undefined;
  const resourceSize =
    resolved && typeof resolved !== 'string' && resolved.kind === 'lazy'
      ? resolved.size
      : sourceEntry instanceof Uint8Array
      ? sourceEntry.byteLength
      : undefined;
  const previewMetadata: PresentationImagePreviewMetadata = {
    sourceKind: detectPresentationImageSourceKind(
      target,
      Boolean(svgBlip),
      resolved && typeof resolved !== 'string' && resolved.kind === 'lazy'
        ? resolved.mimeType
        : undefined,
    ),
    mimeType:
      resolved && typeof resolved !== 'string' && resolved.kind === 'lazy'
        ? resolved.mimeType
        : target
        ? getPresentationMediaMimeType(target)
        : undefined,
    objectName: attr(objectProperties, 'name') ?? undefined,
    objectDescription:
      attr(objectProperties, 'descr') ??
      attr(objectProperties, 'title') ??
      target?.split('/').pop() ??
      undefined,
    resourceKey: target,
    resourceReuseCount: target
      ? packageState.mediaUseCounts?.[target]
      : undefined,
    resourceSize,
  };
  const crop: ImageCrop | undefined = srcRect
    ? {
        left: attr(srcRect, 'l')
          ? Number(attr(srcRect, 'l')) / 100000
          : undefined,
        top: attr(srcRect, 't')
          ? Number(attr(srcRect, 't')) / 100000
          : undefined,
        right: attr(srcRect, 'r')
          ? Number(attr(srcRect, 'r')) / 100000
          : undefined,
        bottom: attr(srcRect, 'b')
          ? Number(attr(srcRect, 'b')) / 100000
          : undefined,
      }
    : undefined;

  return {
    id: `image-${index}`,
    type: 'image',
    previewable,
    previewMetadata,
    x: emuValue(childByLocalName(xfrm, 'off'), 'x') ?? 0,
    y: emuValue(childByLocalName(xfrm, 'off'), 'y') ?? 0,
    width: emuValue(childByLocalName(xfrm, 'ext'), 'cx') ?? 0,
    height: emuValue(childByLocalName(xfrm, 'ext'), 'cy') ?? 0,
    src: resolved || '',
    rotate: attr(xfrm, 'rot') ? Number(attr(xfrm, 'rot')) / 60000 : undefined,
    flipH: attr(xfrm, 'flipH') === '1',
    flipV: attr(xfrm, 'flipV') === '1',
    crop,
    alt: target?.split('/').pop(),
  };
}

function parseWpsWebExtensionChart(
  node: Element,
  index: number,
  packageState: PackageState,
  rels: Record<string, OfficeRelationship>,
): ChartElement | undefined {
  // 关系和位置属于 PPTX 包装层，WPS JSON 到图表模型的转换由共享适配器负责。
  const webExtensionRef = descendantByLocalName(node, 'webExtensionRef');
  const relId = attr(webExtensionRef, 'r:id') ?? attr(webExtensionRef, 'id');
  const target = relId ? rels[relId]?.target : undefined;
  const webExtensionPath = resolveXmlTarget(target, packageState);
  const xml = webExtensionPath
    ? (packageState.entries.get(webExtensionPath) as string | undefined)
    : undefined;
  if (!xml || !webExtensionPath) return undefined;

  const doc = parseXml(xml);
  const snapshotSrc = resolveWebExtensionSnapshot(
    doc,
    webExtensionPath,
    packageState,
  );
  const chart = parseWpsWebExtensionChartModel(
    doc.documentElement,
    typeof snapshotSrc === 'string' ? snapshotSrc : undefined,
  );
  if (!chart) return undefined;

  const xfrm = childByLocalName(childByLocalName(node, 'spPr') ?? node, 'xfrm');
  const off = childByLocalName(xfrm, 'off');
  const ext = childByLocalName(xfrm, 'ext');

  return {
    id: `chart-${index}`,
    type: 'chart',
    chart,
    chartId: relId,
    chartPath: webExtensionPath,
    snapshotSource:
      snapshotSrc && typeof snapshotSrc !== 'string' ? snapshotSrc : undefined,
    x: emuValue(off, 'x') ?? 0,
    y: emuValue(off, 'y') ?? 0,
    width: emuValue(ext, 'cx') ?? 0,
    height: emuValue(ext, 'cy') ?? 0,
  };
}

/** 解析图表部件关系中声明的专属主题覆盖。 */
function resolveChartTheme(
  chartPath: string,
  theme: ThemeModel,
  packageState: PackageState,
): OfficeTheme {
  const relationships =
    packageState.relationships[getOfficePartRelationshipsPath(chartPath)] ?? {};
  const overridePath = Object.values(relationships).find((relationship) =>
    relationship.type?.toLowerCase().endsWith('/themeoverride'),
  )?.target;
  const overrideXml = overridePath
    ? packageState.entries.get(overridePath)
    : undefined;
  return typeof overrideXml === 'string'
    ? readOfficeTheme(overrideXml, theme)
    : theme;
}

function parseChartElement(
  node: Element,
  index: number,
  theme: ThemeModel,
  packageState: PackageState,
  rels: Record<string, OfficeRelationship>,
): ChartElement | undefined {
  const chartNode = descendantByLocalName(node, 'chart');
  const relId = attr(chartNode, 'r:id') ?? attr(chartNode, 'id');
  const target = relId ? rels[relId]?.target : undefined;
  const chartPath = resolveXmlTarget(target, packageState);
  const xml = chartPath
    ? (packageState.entries.get(chartPath) as string | undefined)
    : undefined;
  if (!chartPath || !xml) return undefined;

  const chart = parseOfficeChartXml(
    xml,
    resolveChartTheme(chartPath, theme, packageState),
  );
  const xfrm = childByLocalName(node, 'xfrm');
  const off = childByLocalName(xfrm, 'off');
  const ext = childByLocalName(xfrm, 'ext');

  return {
    id: `chart-${index}`,
    type: 'chart',
    chart,
    chartId: relId,
    chartPath,
    x: emuValue(off, 'x') ?? 0,
    y: emuValue(off, 'y') ?? 0,
    width: emuValue(ext, 'cx') ?? 0,
    height: emuValue(ext, 'cy') ?? 0,
  };
}

function parseTableCellText(
  cellNode: Element,
  theme: ThemeModel,
  relationships: Record<string, OfficeRelationship>,
  slideTargets?: PptxSlideTargetMap,
) {
  const txBody = childByLocalName(cellNode, 'txBody');
  const paragraphs = childrenByLocalName(txBody, 'p').map((paragraphNode) => {
    const paragraphProps = childByLocalName(paragraphNode, 'pPr');
    const paragraphStyle = readParagraphLevelStyle(paragraphProps, theme);
    const defaultRunStyle = mergeTextStyles(
      readDefaultRunStyle(childByLocalName(paragraphProps, 'defRPr'), theme),
      readDefaultRunStyle(childByLocalName(paragraphNode, 'endParaRPr'), theme),
    );
    const runs: TextRun[] = [];

    Array.from(paragraphNode.children).forEach((child) => {
      if (child.localName === 'r' || child.localName === 'fld') {
        const runProps = childByLocalName(child, 'rPr');
        runs.push({
          text:
            textContent(childByLocalName(child, 't')) ||
            child.textContent ||
            '',
          hyperlink: parsePptxHyperlink(runProps, relationships, slideTargets),
          style: mergeTextStyles(
            defaultRunStyle,
            readDefaultRunStyle(runProps, theme),
          ),
        });
        return;
      }

      if (child.localName === 'br') {
        const runProps = childByLocalName(child, 'rPr');
        runs.push({
          text: '\n',
          style: mergeTextStyles(
            defaultRunStyle,
            readDefaultRunStyle(runProps, theme),
          ),
        });
      }
    });

    if (!runs.length && paragraphNode.textContent) {
      runs.push({ text: paragraphNode.textContent, style: defaultRunStyle });
    }

    return {
      runs,
      style: paragraphStyle,
      bullet: paragraphStyle.bullet,
    };
  });

  const text = paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n');
  const firstRunStyle = paragraphs
    .flatMap((paragraph) => paragraph.runs)
    .find(Boolean)?.style;

  return { text, paragraphs, firstRunStyle };
}

function parseTableBorder(tcPr: Element | null, theme: ThemeModel) {
  const line =
    childByLocalName(tcPr, 'ln') ??
    childByLocalName(tcPr, 'lnL') ??
    childByLocalName(tcPr, 'lnR') ??
    childByLocalName(tcPr, 'lnT') ??
    childByLocalName(tcPr, 'lnB');
  if (!line) return {};
  if (childByLocalName(line, 'noFill')) {
    return {
      borderColor: null,
    };
  }
  const fill = childByLocalName(line, 'solidFill') ?? line;
  return {
    borderColor: parseColorNode(fill, theme),
    borderOpacity: parseAlphaNode(fill),
    borderWidth: attr(line, 'w') ? Number(attr(line, 'w')) / 12700 : undefined,
  };
}

function tableFlag(node: Element | null, name: string) {
  return attr(node, name) === '1' || attr(node, name) === 'true';
}

/** 解析并确定 `resolveTableCellStyle` 对应的引用或配置。 */
function resolveTableCellStyle(
  style: TableStyleDefinition | undefined,
  tblPr: Element | null,
  rowIndex: number,
  columnIndex: number,
  rowCount: number,
  columnCount: number,
) {
  if (!style) return {};
  const isFirstRow = tableFlag(tblPr, 'firstRow') && rowIndex === 0;
  const isLastRow = tableFlag(tblPr, 'lastRow') && rowIndex === rowCount - 1;
  const isFirstCol = tableFlag(tblPr, 'firstCol') && columnIndex === 0;
  const isLastCol =
    tableFlag(tblPr, 'lastCol') && columnIndex === columnCount - 1;
  const bandRowOffset = tableFlag(tblPr, 'firstRow') ? 1 : 0;
  const bandColOffset = tableFlag(tblPr, 'firstCol') ? 1 : 0;
  const rowBand =
    tableFlag(tblPr, 'bandRow') && !isFirstRow && !isLastRow
      ? (rowIndex - bandRowOffset) % 2 === 0
        ? style.variants.band1H
        : style.variants.band2H
      : undefined;
  const colBand =
    tableFlag(tblPr, 'bandCol') && !isFirstCol && !isLastCol
      ? (columnIndex - bandColOffset) % 2 === 0
        ? style.variants.band1V
        : style.variants.band2V
      : undefined;

  return mergeTableCellStyle(
    style.variants.wholeTbl,
    rowBand,
    colBand,
    isFirstRow ? style.variants.firstRow : undefined,
    isLastRow ? style.variants.lastRow : undefined,
    isFirstCol ? style.variants.firstCol : undefined,
    isLastCol ? style.variants.lastCol : undefined,
  );
}

function parseTableElement(
  node: Element,
  index: number,
  theme: ThemeModel,
  tableStyles?: TableStyleMap,
  relationships: Record<string, OfficeRelationship> = {},
  slideTargets?: PptxSlideTargetMap,
): TableElement {
  const xfrm = childByLocalName(node, 'xfrm');
  const off = childByLocalName(xfrm, 'off');
  const ext = childByLocalName(xfrm, 'ext');
  const tbl = descendantByLocalName(node, 'tbl');
  const tblPr = childByLocalName(tbl, 'tblPr');
  const styleId = textContent(childByLocalName(tblPr, 'tableStyleId')).trim();
  const tableStyle = styleId ? tableStyles?.[styleId] : undefined;
  const columnWidths = childrenByLocalName(
    childByLocalName(tbl, 'tblGrid'),
    'gridCol',
  ).map((col) => emuValue(col, 'w') ?? 0);
  const rowNodes = childrenByLocalName(tbl, 'tr');
  const rowHeights = rowNodes.map((rowNode) => emuValue(rowNode, 'h') ?? 0);
  const rows: TableCell[][] = rowNodes.map((rowNode, rowIndex) =>
    childrenByLocalName(rowNode, 'tc').map(
      (cellNode, columnIndex): TableCell => {
        const tcPr = childByLocalName(cellNode, 'tcPr');
        const fillNode =
          childByLocalName(tcPr, 'solidFill') ??
          childByLocalName(tcPr, 'gradFill');
        const { text, paragraphs, firstRunStyle } = parseTableCellText(
          cellNode,
          theme,
          relationships,
          slideTargets,
        );
        const explicitBorder = parseTableBorder(tcPr, theme);
        const styled = resolveTableCellStyle(
          tableStyle,
          tblPr,
          rowIndex,
          columnIndex,
          rowNodes.length,
          columnWidths.length,
        );
        const explicitBackgroundColor = childByLocalName(tcPr, 'noFill')
          ? null
          : parseColorNode(fillNode, theme);
        const explicitBackgroundOpacity = childByLocalName(tcPr, 'noFill')
          ? undefined
          : parseAlphaNode(fillNode);
        return {
          text,
          paragraphs,
          style: mergeTextStyles(styled.text, firstRunStyle),
          backgroundColor:
            explicitBackgroundColor !== undefined
              ? explicitBackgroundColor
              : styled.backgroundColor ?? undefined,
          backgroundOpacity:
            explicitBackgroundOpacity !== undefined
              ? explicitBackgroundOpacity
              : styled.backgroundOpacity,
          borderColor:
            explicitBorder.borderColor !== undefined
              ? explicitBorder.borderColor
              : styled.borderColor ?? undefined,
          borderOpacity:
            explicitBorder.borderOpacity !== undefined
              ? explicitBorder.borderOpacity
              : styled.borderOpacity,
          borderWidth:
            explicitBorder.borderWidth !== undefined
              ? explicitBorder.borderWidth
              : styled.borderWidth,
          margins: {
            left: emuValue(tcPr, 'marL'),
            right: emuValue(tcPr, 'marR'),
            top: emuValue(tcPr, 'marT'),
            bottom: emuValue(tcPr, 'marB'),
          },
          verticalAlign:
            attr(tcPr, 'anchor') === 'b'
              ? 'bottom'
              : attr(tcPr, 'anchor') === 'ctr'
              ? 'middle'
              : 'top',
        };
      },
    ),
  );

  return {
    id: `table-${index}`,
    type: 'table',
    x: emuValue(off, 'x') ?? 0,
    y: emuValue(off, 'y') ?? 0,
    width: emuValue(ext, 'cx') ?? 0,
    height: emuValue(ext, 'cy') ?? 0,
    columnWidths,
    rowHeights,
    rows,
  };
}

function parseUnsupportedElement(
  index: number,
  reason: string,
): UnsupportedElement {
  return {
    id: `unsupported-${index}`,
    type: 'unsupported',
    x: 0,
    y: 0,
    width: 120,
    height: 32,
    reason,
  };
}

/** 查找 `findLayoutForSlide` 对应的目标数据。 */
function findLayoutForSlide(
  slideRels: Record<string, OfficeRelationship>,
  layoutDefinitions: LayoutDefinition[],
) {
  const layoutTarget = Object.values(slideRels).find((relationship) =>
    relationship.target.includes('slideLayout'),
  )?.target;
  if (!layoutTarget) return undefined;
  const layoutPath = layoutTarget.startsWith('ppt/')
    ? layoutTarget
    : `ppt/${layoutTarget.replace(/^\.\.\//, '')}`;
  return layoutDefinitions.find(
    (layout) => layout.path === layoutTarget || layout.path === layoutPath,
  );
}

/** 为母版和版式中复用的动态页码创建当前幻灯片专属文本副本。 */
function resolvePptxSlideFields(elements: SlideElement[], slideNumber: number) {
  return elements.map((element) => {
    if (element.type !== 'text') return element;
    let changed = false;
    const paragraphs = element.paragraphs.map((paragraph) => ({
      ...paragraph,
      runs: paragraph.runs.map((run) => {
        if (run.fieldType?.toLowerCase() !== 'slidenum') return run;
        changed = true;
        return { ...run, text: String(slideNumber) };
      }),
    }));
    return changed ? { ...element, paragraphs } : element;
  });
}

/** 解析单张 PPTX 幻灯片及其关系部件。 */
export function parseSlideXml(
  xml: string,
  index: number,
  width: number,
  height: number,
  packageState: PackageState,
  theme: ThemeModel,
  relPath: string,
  layoutDefinitions: LayoutDefinition[],
  masterDefinitions: MasterDefinition[],
  tableStyles?: TableStyleMap,
  slideTargets?: PptxSlideTargetMap,
  defaultTextStyle?: PlaceholderStyle,
): SlideModel {
  const doc = parseXml(xml);
  const slide = doc.documentElement;
  const cSld = childByLocalName(slide, 'cSld');
  const spTree = childByLocalName(cSld, 'spTree');
  const bg = childByLocalName(cSld, 'bg');
  const slideRels = packageState.relationships[relPath] ?? {};
  const layout = findLayoutForSlide(slideRels, layoutDefinitions);
  const master = layout
    ? masterDefinitions.find((item) => item.path === layout.masterPath)
    : masterDefinitions[0];
  const slideBackground = readPptxSlideBackground(
    bg,
    theme,
    packageState,
    slideRels,
  );
  const background =
    slideBackground?.fill || slideBackground?.imageRef
      ? slideBackground
      : layout?.background?.fill || layout?.background?.imageRef
      ? layout.background
      : master?.background;

  // 幻灯片或版式关闭母版图形时，仅跳过母版视觉元素；母版背景和文字样式仍参与继承。
  const showMasterShapes =
    attr(slide, 'showMasterSp') !== '0' && layout?.showMasterShapes !== false;

  const elements: SlideElement[] = [
    ...(showMasterShapes
      ? resolvePptxSlideFields(master?.elements ?? [], index)
      : []),
    ...resolvePptxSlideFields(layout?.elements ?? [], index),
  ];
  const placeholderStyles = buildPlaceholderStyles(
    master,
    layout,
    defaultTextStyle,
  );

  elements.push(
    ...parsePptxVisualTree(
      spTree,
      theme,
      packageState,
      slideRels,
      `slide-${index}`,
      placeholderStyles,
      tableStyles,
      true,
      { slideNumber: index },
      slideTargets,
      defaultTextStyle,
    ),
  );
  const elementsWithPreviewMetadata = annotatePresentationImageReuse(elements);

  const commentResult = parsePptxComments(
    packageState,
    relPath,
    `slide-${index}`,
    index - 1,
    width,
    height,
    elementsWithPreviewMetadata,
  );
  const transitionResult = parsePptxTransition(
    childByLocalName(slide, 'transition'),
    index - 1,
  );
  const warnings = [
    ...commentResult.warnings,
    ...(transitionResult.warning ? [transitionResult.warning] : []),
  ];

  return {
    id: `slide-${index}`,
    index,
    width,
    height,
    background,
    speakerNotes: parsePptxSpeakerNotes(
      packageState.entries,
      packageState.relationships,
      relPath,
    ),
    annotations: commentResult.annotations,
    transition: transitionResult.transition,
    warnings: warnings.length ? warnings : undefined,
    elements: elementsWithPreviewMetadata,
  };
}
