import { parseOfficeChartXml } from '../../shared/ooxml/charts';
import { resolvePackageMediaRef } from '../../shared/ooxml/media';
import type { OfficeTheme } from '../../shared/ooxml/theme';
import { emuToPx } from '../../shared/ooxml/units';
import { parseWpsWebExtensionChartModel } from '../../shared/ooxml/wpsChart';
import {
  attr,
  childByLocalName,
  descendantByLocalName,
  descendantsByLocalName,
  matchesLocalName,
  parseXml,
  textContent,
} from '../../shared/ooxml/xml';
import {
  DEFAULT_DOCX_PAGE,
  normalizeCssColor,
  readCssDeclaration,
  readCssPosition,
  readCssSize,
  readDrawingColor,
  vmlUnitToPx,
  type DocxPackageState,
  type DocxParseContext,
  type ReadBlockChildrenOptions,
} from './docxParsingContext';
import { parseDocxDrawingHyperlink } from './parseDocxHyperlink';
import type {
  DocxBlock,
  DocxChartBlock,
  DocxImage,
  DocxInline,
  DocxParagraphBlock,
  DocxPosition,
  DocxShape,
  DocxShapeItem,
} from './types';

/** DOCX 绘图解析流程共享的资源和样式上下文。 */
type ParseContext = DocxParseContext;

/** 读取绘图文本框内块级内容的回调。 */
export type DocxDrawingContentReader = (
  node: Element | null | undefined,
  id: string,
  context: DocxParseContext,
  options?: ReadBlockChildrenOptions,
) => DocxBlock[];

/** DOCX 行内绘图解析器。 */
export type DocxDrawingParser = {
  /** 将绘图、VML 或兼容内容节点转换为行内模型。 */
  parseRunChild(
    node: Element,
    context: DocxParseContext,
  ): DocxInline | undefined;
};

function resolveMediaRef(
  target: string | undefined,
  packageState: DocxPackageState,
) {
  return resolvePackageMediaRef(
    target,
    packageState.mediaByPath,
    packageState.mediaByName,
    'word',
  );
}

/** 解析并确定 `resolveXmlTarget` 对应的引用或配置。 */
function resolveXmlTarget(
  target: string | undefined,
  packageState: DocxPackageState,
) {
  if (!target) return undefined;
  const normalized = target.replace(/^\.\.\//, '');
  return packageState.entries.get(normalized) ? normalized : target;
}

function readDrawingAnchorPosition(
  node: Element,
  documentGridLineHeight?: number,
) {
  const anchor = descendantByLocalName(node, 'anchor');
  if (!anchor) return undefined;

  const positionH = childByLocalName(anchor, 'positionH');
  const positionV = childByLocalName(anchor, 'positionV');
  const relativeFromV = attr(positionV, 'relativeFrom');
  const affectsTextFlow = [
    'wrapSquare',
    'wrapTight',
    'wrapThrough',
    'wrapTopAndBottom',
  ].some((name) => Boolean(childByLocalName(anchor, name)));
  const left = emuToPx(
    Number(textContent(childByLocalName(positionH, 'posOffset')).trim()),
  );
  const top =
    emuToPx(
      Number(textContent(childByLocalName(positionV, 'posOffset')).trim()),
    ) +
    (relativeFromV === 'paragraph' && affectsTextFlow
      ? documentGridLineHeight ?? 0
      : 0);
  if (!Number.isFinite(left) || !Number.isFinite(top)) return undefined;

  const relativeHeight = Number(attr(anchor, 'relativeHeight'));
  const rotation = Number(attr(anchor, 'rotation'));

  return {
    left: Math.round(left),
    top: Math.round(top),
    relativeFromH: attr(
      positionH,
      'relativeFrom',
    ) as DocxPosition['relativeFromH'],
    relativeFromV: attr(
      positionV,
      'relativeFrom',
    ) as DocxPosition['relativeFromV'],
    zIndex: Number.isFinite(relativeHeight) ? relativeHeight : undefined,
    behindDoc: attr(anchor, 'behindDoc') === '1',
    rotation:
      Number.isFinite(rotation) && rotation !== 0
        ? rotation / 60000
        : undefined,
    flipH: attr(anchor, 'flipH') === '1' || undefined,
    flipV: attr(anchor, 'flipV') === '1' || undefined,
  };
}

function parseChartElement(
  node: Element,
  context: ParseContext,
): DocxChartBlock | undefined {
  const chartNode = descendantByLocalName(node, 'chart');
  const relId = attr(chartNode, 'r:id') ?? attr(chartNode, 'id');
  const target = relId ? context.documentRels[relId]?.target : undefined;
  const chartPath = resolveXmlTarget(target, context.packageState);
  const xml = chartPath
    ? (context.packageState.entries.get(chartPath) as string | undefined)
    : undefined;
  if (!xml) return undefined;

  const chart = parseOfficeChartXml(xml, context.theme);
  const extent =
    descendantByLocalName(node, 'extent') ??
    descendantByLocalName(node, 'xfrm');
  const width = Math.max(
    160,
    Math.round(emuToPx(Number(attr(extent, 'cx') ?? 0)) || 320),
  );
  const height = Math.max(
    120,
    Math.round(emuToPx(Number(attr(extent, 'cy') ?? 0)) || 220),
  );
  context.chartIndex += 1;
  return {
    id: `docx-chart-${context.chartIndex}`,
    type: 'chart',
    chart,
    width,
    height,
  };
}

/** 解析并确定 `resolveWebExtensionSnapshot` 对应的引用或配置。 */
function resolveWebExtensionSnapshot(
  doc: XMLDocument,
  webExtensionPath: string,
  context: ParseContext,
) {
  const snapshot = descendantByLocalName(doc.documentElement, 'snapshot');
  const embed = attr(snapshot, 'r:embed') ?? attr(snapshot, 'embed');
  const relsPath = webExtensionPath
    .replace(/^word\/webExtensions\//, 'word/webExtensions/_rels/')
    .concat('.rels');
  const target = embed
    ? context.packageState.relationships[relsPath]?.[embed]?.target
    : undefined;
  return resolveMediaRef(target, context.packageState);
}

function parseWpsWebExtensionChart(
  node: Element,
  context: ParseContext,
): DocxChartBlock | undefined {
  // 关系和尺寸属于 DOCX 包装层，WPS JSON 到图表模型的转换由共享适配器负责。
  const webExtensionRef = descendantByLocalName(node, 'webExtensionRef');
  const relId = attr(webExtensionRef, 'r:id') ?? attr(webExtensionRef, 'id');
  const target = relId ? context.documentRels[relId]?.target : undefined;
  const webExtensionPath = resolveXmlTarget(target, context.packageState);
  const xml = webExtensionPath
    ? (context.packageState.entries.get(webExtensionPath) as string | undefined)
    : undefined;
  if (!xml || !webExtensionPath) return undefined;

  const doc = parseXml(xml);
  const snapshotSrc = resolveWebExtensionSnapshot(
    doc,
    webExtensionPath,
    context,
  );
  const chart = parseWpsWebExtensionChartModel(
    doc.documentElement,
    snapshotSrc?.kind === 'url' ? snapshotSrc.url : undefined,
  );
  if (!chart) return undefined;

  const extent =
    descendantByLocalName(node, 'extent') ??
    descendantByLocalName(node, 'xfrm');
  const width = Math.max(
    160,
    Math.round(emuToPx(Number(attr(extent, 'cx') ?? 0)) || 320),
  );
  const height = Math.max(
    120,
    Math.round(emuToPx(Number(attr(extent, 'cy') ?? 0)) || 220),
  );
  context.chartIndex += 1;

  return {
    id: `docx-chart-${context.chartIndex}`,
    type: 'chart',
    chart,
    snapshotSource: snapshotSrc,
    width,
    height,
  };
}

function readTopLevelDrawingGraphicData(drawingNode: Element) {
  const drawingContainer =
    childByLocalName(drawingNode, 'anchor') ??
    childByLocalName(drawingNode, 'inline');
  return childByLocalName(
    childByLocalName(drawingContainer, 'graphic'),
    'graphicData',
  );
}

function isDirectDrawingPicture(drawingNode: Element) {
  const graphicData = readTopLevelDrawingGraphicData(drawingNode);
  return (
    attr(graphicData, 'uri') ===
      'http://schemas.openxmlformats.org/drawingml/2006/picture' &&
    Boolean(childByLocalName(graphicData, 'pic'))
  );
}

/** 判断绘图是否包含不能退化为单张图片的组合内容。 */
function hasCompoundDrawingContent(drawingNode: Element) {
  const graphicData = readTopLevelDrawingGraphicData(drawingNode);
  const group = descendantByLocalName(graphicData, 'wgp');
  if (!group) return Boolean(childByLocalName(graphicData, 'wsp'));

  const drawableChildren = Array.from(group.children).filter(
    (child) => matchesLocalName(child, 'wsp') || matchesLocalName(child, 'pic'),
  );
  return (
    drawableChildren.some((child) => matchesLocalName(child, 'wsp')) ||
    drawableChildren.length > 1
  );
}

/** 读取 DrawingML 图片的颜色替换效果；浏览器原生图片无法直接表达该效果。 */
function readDrawingImageColorChange(
  drawingNode: Element,
): DocxImage['colorChange'] | undefined {
  const blip = descendantByLocalName(drawingNode, 'blip');
  const colorChange = childByLocalName(blip, 'clrChange');
  if (!colorChange) return undefined;

  const fromNode = childByLocalName(
    childByLocalName(colorChange, 'clrFrom'),
    'srgbClr',
  );
  const toNode = childByLocalName(
    childByLocalName(colorChange, 'clrTo'),
    'srgbClr',
  );
  const from = normalizeCssColor(attr(fromNode, 'val'));
  const to = normalizeCssColor(attr(toNode, 'val'));
  if (!from || !to) return undefined;

  const alphaValues = Array.from(toNode?.children ?? []).filter((child) =>
    matchesLocalName(child, 'alpha'),
  );
  const rawAlpha = Number(
    attr(alphaValues[alphaValues.length - 1], 'val') ?? 100000,
  );
  return {
    from,
    to,
    alpha: Number.isFinite(rawAlpha)
      ? Math.max(0, Math.min(1, rawAlpha / 100000))
      : 1,
    useAlpha: attr(colorChange, 'useA') === '1' || undefined,
  };
}

/** 读取 DrawingML 图片四边以十万分比声明的裁剪范围。 */
function readDrawingImageCrop(
  drawingNode: Element,
): DocxImage['crop'] | undefined {
  const picture = descendantByLocalName(drawingNode, 'pic');
  const blipFill = childByLocalName(picture, 'blipFill');
  const srcRect = childByLocalName(blipFill, 'srcRect');
  if (!srcRect) return undefined;

  const readSide = (name: 'l' | 't' | 'r' | 'b') => {
    const value = Number(attr(srcRect, name) ?? 0) / 100000;
    return Number.isFinite(value) ? Math.min(0.99, Math.max(0, value)) : 0;
  };
  return {
    left: readSide('l'),
    top: readSide('t'),
    right: readSide('r'),
    bottom: readSide('b'),
  };
}

function parseDrawingImage(
  drawingNode: Element,
  context: ParseContext,
): DocxImage | undefined {
  const blip = descendantByLocalName(drawingNode, 'blip');
  const embed = attr(blip, 'r:embed') ?? attr(blip, 'embed');
  const target = embed ? context.documentRels[embed]?.target : undefined;
  const src = resolveMediaRef(target, context.packageState);
  if (!src) return undefined;

  const extent = descendantByLocalName(drawingNode, 'extent');
  const docPr = descendantByLocalName(drawingNode, 'docPr');
  const width = Math.max(
    1,
    Math.round(emuToPx(Number(attr(extent, 'cx') ?? 0))),
  );
  const height = Math.max(
    1,
    Math.round(emuToPx(Number(attr(extent, 'cy') ?? 0))),
  );
  const name = attr(docPr, 'name');
  const anchorPosition = readDrawingAnchorPosition(
    drawingNode,
    context.documentGridLineHeight,
  );
  const imageTransform = readDrawingImageTransform(drawingNode);
  const position = anchorPosition
    ? {
        ...anchorPosition,
        flipH: imageTransform.flipH ?? anchorPosition.flipH,
        flipV: imageTransform.flipV ?? anchorPosition.flipV,
      }
    : undefined;
  const image: DocxImage = {
    id: `docx-image-${context.imageIndex + 1}`,
    name,
    alt: attr(docPr, 'descr') ?? name,
    src,
    width,
    height,
    crop: readDrawingImageCrop(drawingNode),
    colorChange: readDrawingImageColorChange(drawingNode),
    position,
    hyperlink: parseDocxDrawingHyperlink(
      docPr ?? drawingNode,
      context.documentRels,
    ),
  };
  context.imageIndex += 1;
  context.images.push(image);
  return image;
}

function untrackParsedImage(context: ParseContext, image: DocxImage) {
  if (context.images[context.images.length - 1]?.id !== image.id) return;
  context.images.pop();
  context.imageIndex = Math.max(0, context.imageIndex - 1);
}

function isLikelyPageSizedNestedImage(image: DocxImage) {
  return (
    image.width >= DEFAULT_DOCX_PAGE.width * 0.75 &&
    image.height >= DEFAULT_DOCX_PAGE.minHeight * 0.7
  );
}

function parseAlternateContentImage(
  drawingNode: Element,
  context: ParseContext,
) {
  // 图片与文本、多个图片组成的组必须交给形状解析器，否则只会保留首张图片。
  if (hasCompoundDrawingContent(drawingNode)) return undefined;
  const image = parseDrawingImage(drawingNode, context);
  if (!image) return undefined;
  if (
    isDirectDrawingPicture(drawingNode) ||
    !isLikelyPageSizedNestedImage(image)
  )
    return image;
  untrackParsedImage(context, image);
  return undefined;
}

function readDrawingImageTransform(
  drawingNode: Element,
): Pick<DocxPosition, 'flipH' | 'flipV'> {
  const picture = descendantByLocalName(drawingNode, 'pic');
  const xfrm = childByLocalName(childByLocalName(picture, 'spPr'), 'xfrm');
  return {
    flipH: attr(xfrm, 'flipH') === '1' || undefined,
    flipV: attr(xfrm, 'flipV') === '1' || undefined,
  };
}

function readDrawingNoFill(node: Element | null | undefined) {
  return Boolean(childByLocalName(node, 'noFill'));
}

function readDrawingTransparentFill(node: Element | null | undefined) {
  const solidFill =
    childByLocalName(node, 'solidFill') ??
    (matchesLocalName(node, 'solidFill') ? node : null);
  const colorNodes = ['srgbClr', 'schemeClr', 'sysClr']
    .map((name) => childByLocalName(solidFill, name))
    .filter((colorNode): colorNode is Element => Boolean(colorNode));
  return (
    colorNodes.length > 0 &&
    colorNodes.every(
      (colorNode) => attr(childByLocalName(colorNode, 'alpha'), 'val') === '0',
    )
  );
}

function parseDrawingLineStyle(
  spPr: Element | null | undefined,
  theme: OfficeTheme,
) {
  const line = childByLocalName(spPr, 'ln');
  if (!line || readDrawingNoFill(line) || readDrawingTransparentFill(line))
    return {};
  const width = emuToPx(Number(attr(line, 'w') ?? 0)) || 1;
  const color = readDrawingColor(line, theme) ?? '#000';
  const dash = attr(childByLocalName(line, 'prstDash'), 'val');
  const borderStyle =
    dash && dash !== 'solid'
      ? dash.toLowerCase().includes('dot')
        ? 'dotted'
        : 'dashed'
      : 'solid';
  const strokeDasharray =
    dash && dash !== 'solid' ? `${width * 3} ${width}` : undefined;
  return {
    border: `${width}px ${borderStyle} ${color}`,
    strokeColor: color,
    strokeWidth: width,
    strokeDasharray,
  };
}

function parseDrawingFillColor(
  spPr: Element | null | undefined,
  theme: OfficeTheme,
) {
  if (readDrawingNoFill(spPr)) return undefined;
  return readDrawingColor(spPr, theme);
}

function parseDrawingFillImage(
  spPr: Element | null | undefined,
  context: ParseContext,
) {
  const blipFill = childByLocalName(spPr, 'blipFill');
  const blip = childByLocalName(blipFill, 'blip');
  const embed = attr(blip, 'r:embed') ?? attr(blip, 'embed');
  const target = embed ? context.documentRels[embed]?.target : undefined;
  return resolveMediaRef(target, context.packageState);
}

/** 二维仿射矩阵，用于把嵌套组合图形的局部坐标映射到最终像素坐标。 */
type DrawingMatrix = {
  /** 水平轴的水平分量。 */
  a: number;
  /** 水平轴的垂直分量。 */
  b: number;
  /** 垂直轴的水平分量。 */
  c: number;
  /** 垂直轴的垂直分量。 */
  d: number;
  /** 水平位移。 */
  e: number;
  /** 垂直位移。 */
  f: number;
};

const IDENTITY_DRAWING_MATRIX: DrawingMatrix = {
  a: 1,
  b: 0,
  c: 0,
  d: 1,
  e: 0,
  f: 0,
};

/** 组合两个坐标变换，结果会先应用 child，再应用 parent。 */
function multiplyDrawingMatrices(
  parent: DrawingMatrix,
  child: DrawingMatrix,
): DrawingMatrix {
  return {
    a: parent.a * child.a + parent.c * child.b,
    b: parent.b * child.a + parent.d * child.b,
    c: parent.a * child.c + parent.c * child.d,
    d: parent.b * child.c + parent.d * child.d,
    e: parent.a * child.e + parent.c * child.f + parent.e,
    f: parent.b * child.e + parent.d * child.f + parent.f,
  };
}

function createDrawingTranslationMatrix(x: number, y: number): DrawingMatrix {
  return { ...IDENTITY_DRAWING_MATRIX, e: x, f: y };
}

function createDrawingScaleMatrix(x: number, y: number): DrawingMatrix {
  return { ...IDENTITY_DRAWING_MATRIX, a: x, d: y };
}

function createDrawingRotationMatrix(degrees: number): DrawingMatrix {
  const radians = (degrees * Math.PI) / 180;
  const cosine = Math.cos(radians);
  const sine = Math.sin(radians);
  return {
    a: cosine,
    b: sine,
    c: -sine,
    d: cosine,
    e: 0,
    f: 0,
  };
}

function applyDrawingMatrix(matrix: DrawingMatrix, x: number, y: number) {
  return {
    x: matrix.a * x + matrix.c * y + matrix.e,
    y: matrix.b * x + matrix.d * y + matrix.f,
  };
}

function readFiniteNumber(value: string | undefined, fallback = 0) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

/** 读取组合容器从子坐标系到父坐标系的完整变换。 */
function readDrawingGroupMatrix(groupNode: Element): DrawingMatrix {
  const groupProperties = childByLocalName(groupNode, 'grpSpPr');
  const xfrm = childByLocalName(groupProperties, 'xfrm');
  if (!xfrm) return IDENTITY_DRAWING_MATRIX;

  const off = childByLocalName(xfrm, 'off');
  const ext = childByLocalName(xfrm, 'ext');
  const childOff = childByLocalName(xfrm, 'chOff');
  const childExt = childByLocalName(xfrm, 'chExt');
  const left = readFiniteNumber(attr(off, 'x'));
  const top = readFiniteNumber(attr(off, 'y'));
  const width = readFiniteNumber(attr(ext, 'cx'));
  const height = readFiniteNumber(attr(ext, 'cy'));
  const childLeft = readFiniteNumber(attr(childOff, 'x'));
  const childTop = readFiniteNumber(attr(childOff, 'y'));
  const childWidth = readFiniteNumber(attr(childExt, 'cx'), width || 1);
  const childHeight = readFiniteNumber(attr(childExt, 'cy'), height || 1);
  const scaleX = childWidth > 0 ? width / childWidth : 1;
  const scaleY = childHeight > 0 ? height / childHeight : 1;

  let matrix = multiplyDrawingMatrices(
    createDrawingTranslationMatrix(left, top),
    multiplyDrawingMatrices(
      createDrawingScaleMatrix(scaleX, scaleY),
      createDrawingTranslationMatrix(-childLeft, -childTop),
    ),
  );

  const rawRotation = readFiniteNumber(attr(xfrm, 'rot'));
  const rotation = rawRotation / 60000;
  const flipX = attr(xfrm, 'flipH') === '1' ? -1 : 1;
  const flipY = attr(xfrm, 'flipV') === '1' ? -1 : 1;
  if (rotation || flipX < 0 || flipY < 0) {
    const centerX = left + width / 2;
    const centerY = top + height / 2;
    const aroundCenter = multiplyDrawingMatrices(
      createDrawingTranslationMatrix(centerX, centerY),
      multiplyDrawingMatrices(
        createDrawingRotationMatrix(rotation),
        multiplyDrawingMatrices(
          createDrawingScaleMatrix(flipX, flipY),
          createDrawingTranslationMatrix(-centerX, -centerY),
        ),
      ),
    );
    matrix = multiplyDrawingMatrices(aroundCenter, matrix);
  }

  return matrix;
}

/** 将顶层组合图形的 EMU 坐标映射到锚点提供的像素视口。 */
function readWpgRootMatrix(
  groupNode: Element,
  width: number,
  height: number,
): DrawingMatrix {
  const xfrm = childByLocalName(childByLocalName(groupNode, 'grpSpPr'), 'xfrm');
  const ext = childByLocalName(xfrm, 'ext');
  const rawWidth = readFiniteNumber(attr(ext, 'cx'));
  const rawHeight = readFiniteNumber(attr(ext, 'cy'));
  const viewportMatrix = createDrawingScaleMatrix(
    rawWidth > 0 ? width / rawWidth : emuToPx(1),
    rawHeight > 0 ? height / rawHeight : emuToPx(1),
  );
  return multiplyDrawingMatrices(
    viewportMatrix,
    readDrawingGroupMatrix(groupNode),
  );
}

function readDrawingItemSize(
  node: Element | null | undefined,
  matrix: DrawingMatrix,
) {
  const xfrm = childByLocalName(node, 'xfrm');
  const off = childByLocalName(xfrm, 'off');
  const ext = childByLocalName(xfrm, 'ext');
  const rawLeft = readFiniteNumber(attr(off, 'x'));
  const rawTop = readFiniteNumber(attr(off, 'y'));
  const rawWidth = readFiniteNumber(attr(ext, 'cx'));
  const rawHeight = readFiniteNumber(attr(ext, 'cy'));
  const center = applyDrawingMatrix(
    matrix,
    rawLeft + rawWidth / 2,
    rawTop + rawHeight / 2,
  );
  const width = Math.hypot(matrix.a * rawWidth, matrix.b * rawWidth);
  const height = Math.hypot(matrix.c * rawHeight, matrix.d * rawHeight);
  return {
    left: center.x - width / 2,
    top: center.y - height / 2,
    width,
    height,
  };
}

/** 合并子图形自身与外层组合容器的旋转、翻转。 */
function readDrawingItemTransform(
  node: Element | null | undefined,
  matrix: DrawingMatrix,
): Pick<DocxShapeItem, 'rotation' | 'flipH' | 'flipV'> {
  const xfrm = childByLocalName(node, 'xfrm');
  const rawRotation = readFiniteNumber(attr(xfrm, 'rot')) / 60000;
  const groupRotation = (Math.atan2(matrix.b, matrix.a) * 180) / Math.PI;
  const determinant = matrix.a * matrix.d - matrix.b * matrix.c;
  const groupFlipV = determinant < 0;
  const rotation = groupRotation + (groupFlipV ? -rawRotation : rawRotation);
  return {
    rotation: Math.abs(rotation) > 0.0001 ? rotation : undefined,
    flipH: attr(xfrm, 'flipH') === '1' || undefined,
    flipV: (attr(xfrm, 'flipV') === '1') !== groupFlipV || undefined,
  };
}

/** 递归收集组合图形中的可绘制子项，并保留每一级坐标变换。 */
function collectWpgDrawables(
  groupNode: Element,
  parentMatrix: DrawingMatrix,
  result: Array<{ node: Element; matrix: DrawingMatrix }> = [],
) {
  Array.from(groupNode.children).forEach((child) => {
    if (matchesLocalName(child, 'wsp') || matchesLocalName(child, 'pic')) {
      result.push({ node: child, matrix: parentMatrix });
      return;
    }
    if (matchesLocalName(child, 'grpSp')) {
      collectWpgDrawables(
        child,
        multiplyDrawingMatrices(parentMatrix, readDrawingGroupMatrix(child)),
        result,
      );
    }
  });
  return result;
}

function readDrawingShapeKind(
  spPr: Element | null | undefined,
): DocxShapeItem['kind'] {
  const geometry = childByLocalName(spPr, 'prstGeom');
  const preset = attr(geometry, 'prst');
  if (preset === 'line') return 'line';
  if (preset === 'ellipse') return 'ellipse';
  if (
    preset === 'star5' ||
    preset === 'moon' ||
    preset === 'cloud' ||
    preset === 'heart' ||
    preset === 'horizontalScroll'
  )
    return 'path';
  if (childByLocalName(spPr, 'custGeom')) return 'path';
  return 'rect';
}

function readDrawingShapePreset(spPr: Element | null | undefined) {
  return attr(childByLocalName(spPr, 'prstGeom'), 'prst');
}

/** DrawingML 圆角矩形未声明调整值时使用的标准默认比例。 */
const DEFAULT_DRAWING_ROUNDED_RECTANGLE_RATIO = 16667 / 100000;

/** 读取预设几何的首个调整比例，并约束到该预设允许的范围。 */
function readDrawingPresetAdjustmentRatio(
  geometry: Element | null | undefined,
  defaultRatio: number,
  maximumRatio: number,
) {
  const adjustment = Array.from(
    childByLocalName(geometry, 'avLst')?.children ?? [],
  ).find(
    (child) => matchesLocalName(child, 'gd') && attr(child, 'name') === 'adj',
  );
  const formula = attr(adjustment, 'fmla')?.trim();
  const value = formula?.match(/^val\s+(-?\d+(?:\.\d+)?)$/)?.[1];
  const ratio = Number(value) / 100000;
  if (!Number.isFinite(ratio)) return defaultRatio;
  return Math.min(maximumRatio, Math.max(0, ratio));
}

function readDrawingRoundedRectangleRatio(
  geometry: Element | null | undefined,
) {
  return readDrawingPresetAdjustmentRatio(
    geometry,
    DEFAULT_DRAWING_ROUNDED_RECTANGLE_RATIO,
    0.5,
  );
}

function readDrawingShapeBorderRadius(
  spPr: Element | null | undefined,
  size: {
    /** 对象宽度，单位为标准化渲染像素。 */
    width: number;
    /** 对象高度，单位为标准化渲染像素。 */
    height: number;
  },
) {
  const geometry = childByLocalName(spPr, 'prstGeom');
  const preset = attr(geometry, 'prst');
  if (preset !== 'roundRect' && preset !== 'flowChartAlternateProcess') {
    return undefined;
  }
  return (
    Math.min(size.width, size.height) *
    readDrawingRoundedRectangleRatio(geometry)
  );
}

function readDrawingTextAnchor(
  shapeNode: Element,
): DocxShapeItem['textVerticalAlign'] {
  const anchor = attr(childByLocalName(shapeNode, 'bodyPr'), 'anchor');
  if (anchor === 'ctr') return 'middle';
  if (anchor === 'b') return 'bottom';
  return 'top';
}

function readDrawingTextBehavior(shapeNode: Element) {
  const bodyPr = childByLocalName(shapeNode, 'bodyPr');
  const wrap = attr(bodyPr, 'wrap');
  const verticalOverflow = attr(bodyPr, 'vertOverflow');

  return {
    fitShapeToText: Boolean(childByLocalName(bodyPr, 'spAutoFit')),
    noWrap: wrap === 'none',
    clipVerticalOverflow: verticalOverflow === 'clip',
  };
}

function readDrawingBodyPadding(shapeNode: Element) {
  const bodyPr = childByLocalName(shapeNode, 'bodyPr');
  return {
    paddingTop: vmlUnitToPx(attr(bodyPr, 'tIns')),
    paddingRight: vmlUnitToPx(attr(bodyPr, 'rIns')),
    paddingBottom: vmlUnitToPx(attr(bodyPr, 'bIns')),
    paddingLeft: vmlUnitToPx(attr(bodyPr, 'lIns')),
  };
}

function convertDrawingCustomGeometry(
  spPr: Element | null | undefined,
  width: number,
  height: number,
) {
  const pathNode = descendantByLocalName(
    childByLocalName(spPr, 'custGeom'),
    'path',
  );
  if (!pathNode) return undefined;
  const pathWidth = Number(attr(pathNode, 'w') ?? 0);
  const pathHeight = Number(attr(pathNode, 'h') ?? 0);
  const scaleX =
    Number.isFinite(pathWidth) && pathWidth > 0 ? width / pathWidth : 1;
  const scaleY =
    Number.isFinite(pathHeight) && pathHeight > 0 ? height / pathHeight : 1;
  const commands: string[] = [];

  Array.from(pathNode.children).forEach((child) => {
    if (matchesLocalName(child, 'close')) {
      commands.push('Z');
      return;
    }
    const points = descendantsByLocalName(child, 'pt').map((point) => {
      const x = Number(attr(point, 'x') ?? 0);
      const y = Number(attr(point, 'y') ?? 0);
      return Number.isFinite(x) && Number.isFinite(y)
        ? `${x * scaleX} ${y * scaleY}`
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

  return commands.length ? commands.join(' ') : undefined;
}

/** 把输入格式化为 `formatPathNumber` 返回的展示值。 */
function formatPathNumber(value: number) {
  return Number(value.toFixed(3));
}

/** 按 DrawingML 的 darkenLess 规则将路径填充色降低到 80%。 */
function darkenDrawingPathFillLess(color: string | undefined) {
  const match = color?.match(/^#([0-9a-f]{6})$/i);
  if (!match) return color;
  const channels = match[1].match(/.{2}/g);
  if (!channels) return color;
  return `#${channels
    .map((channel) =>
      Math.round(Number.parseInt(channel, 16) * 0.8)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`;
}

/** 按标准预设几何生成横卷轴的分层填充和轮廓路径。 */
function convertHorizontalScrollPathLayers(
  spPr: Element | null | undefined,
  width: number,
  height: number,
  fillColor: string | undefined,
  stroke: Pick<
    DocxShapeItem,
    'strokeColor' | 'strokeWidth' | 'strokeDasharray'
  >,
): NonNullable<DocxShapeItem['pathLayers']> | undefined {
  if (readDrawingShapePreset(spPr) !== 'horizontalScroll') return undefined;
  const geometry = childByLocalName(spPr, 'prstGeom');
  const shortSide = Math.min(width, height);
  const curl =
    shortSide * readDrawingPresetAdjustmentRatio(geometry, 0.125, 0.25);
  const halfCurl = curl / 2;
  const quarterCurl = curl / 4;
  const x3 = width - curl;
  const x4 = width - halfCurl;
  const y3 = curl + halfCurl;
  const y4 = curl * 2;
  const y5 = height - curl - halfCurl;
  const y6 = height - curl;
  const y7 = height - halfCurl;
  const n = formatPathNumber;
  const basePath = [
    `M ${n(width)} ${n(halfCurl)}`,
    `A ${n(halfCurl)} ${n(halfCurl)} 0 0 1 ${n(x4)} ${n(curl)}`,
    `L ${n(x4)} ${n(halfCurl)}`,
    `A ${n(quarterCurl)} ${n(quarterCurl)} 0 0 1 ${n(x3)} ${n(halfCurl)}`,
    `L ${n(x3)} ${n(curl)} L ${n(halfCurl)} ${n(curl)}`,
    `A ${n(halfCurl)} ${n(halfCurl)} 0 0 0 0 ${n(y3)}`,
    `L 0 ${n(y7)}`,
    `A ${n(halfCurl)} ${n(halfCurl)} 0 0 0 ${n(curl)} ${n(y7)}`,
    `L ${n(curl)} ${n(y6)} L ${n(x4)} ${n(y6)}`,
    `A ${n(halfCurl)} ${n(halfCurl)} 0 0 0 ${n(width)} ${n(y5)} Z`,
    `M ${n(halfCurl)} ${n(y4)}`,
    `A ${n(halfCurl)} ${n(halfCurl)} 0 0 0 ${n(curl)} ${n(y3)}`,
    `A ${n(quarterCurl)} ${n(quarterCurl)} 0 0 0 ${n(halfCurl)} ${n(y3)} Z`,
  ].join(' ');
  const shadedPath = [
    `M ${n(halfCurl)} ${n(y4)}`,
    `A ${n(halfCurl)} ${n(halfCurl)} 0 0 0 ${n(curl)} ${n(y3)}`,
    `A ${n(quarterCurl)} ${n(quarterCurl)} 0 0 0 ${n(halfCurl)} ${n(y3)} Z`,
    `M ${n(x4)} ${n(curl)}`,
    `A ${n(halfCurl)} ${n(halfCurl)} 0 1 0 ${n(x3)} ${n(halfCurl)}`,
    `A ${n(quarterCurl)} ${n(quarterCurl)} 0 0 0 ${n(x4)} ${n(halfCurl)} Z`,
  ].join(' ');
  const outlinePath = [
    `M 0 ${n(y3)}`,
    `A ${n(halfCurl)} ${n(halfCurl)} 0 0 1 ${n(halfCurl)} ${n(y4)}`,
    `L ${n(x3)} ${n(curl)} L ${n(x3)} ${n(halfCurl)}`,
    `A ${n(halfCurl)} ${n(halfCurl)} 0 0 1 ${n(width)} ${n(halfCurl)}`,
    `L ${n(width)} ${n(y5)}`,
    `A ${n(halfCurl)} ${n(halfCurl)} 0 0 1 ${n(x4)} ${n(y6)}`,
    `L ${n(curl)} ${n(y6)} L ${n(curl)} ${n(y7)}`,
    `A ${n(halfCurl)} ${n(halfCurl)} 0 0 1 0 ${n(y7)} Z`,
    `M ${n(x3)} ${n(curl)} L ${n(x4)} ${n(curl)}`,
    `A ${n(halfCurl)} ${n(halfCurl)} 0 0 0 ${n(width)} ${n(halfCurl)}`,
    `M ${n(x4)} ${n(curl)} L ${n(x4)} ${n(halfCurl)}`,
    `A ${n(quarterCurl)} ${n(quarterCurl)} 0 0 1 ${n(x3)} ${n(halfCurl)}`,
    `M ${n(halfCurl)} ${n(y4)} L ${n(halfCurl)} ${n(y3)}`,
    `A ${n(quarterCurl)} ${n(quarterCurl)} 0 0 1 ${n(curl)} ${n(y3)}`,
    `A ${n(halfCurl)} ${n(halfCurl)} 0 0 1 0 ${n(y3)}`,
    `M ${n(curl)} ${n(y3)} L ${n(curl)} ${n(y6)}`,
  ].join(' ');

  return [
    { path: basePath, fillColor },
    { path: shadedPath, fillColor: darkenDrawingPathFillLess(fillColor) },
    { path: outlinePath, ...stroke },
  ];
}

function convertDrawingPresetGeometry(
  spPr: Element | null | undefined,
  width: number,
  height: number,
) {
  const preset = readDrawingShapePreset(spPr);
  if (preset === 'star5') {
    const centerX = width / 2;
    const centerY = height / 2;
    const outerRadiusX = width / 2;
    const outerRadiusY = height / 2;
    const innerRadiusX = outerRadiusX * 0.48;
    const innerRadiusY = outerRadiusY * 0.48;
    const points = Array.from({ length: 10 }, (_, index) => {
      const angle = -Math.PI / 2 + (index * Math.PI) / 5;
      const radiusX = index % 2 === 0 ? outerRadiusX : innerRadiusX;
      const radiusY = index % 2 === 0 ? outerRadiusY : innerRadiusY;
      return `${formatPathNumber(
        centerX + Math.cos(angle) * radiusX,
      )} ${formatPathNumber(centerY + Math.sin(angle) * radiusY)}`;
    });
    return `M ${points[0]} L ${points.slice(1).join(' L ')} Z`;
  }

  if (preset === 'moon') {
    const startX = width * 0.76;
    return [
      `M ${formatPathNumber(startX)} 0`,
      `A ${formatPathNumber(width * 0.7)} ${formatPathNumber(
        height * 0.5,
      )} 0 1 0 ${formatPathNumber(startX)} ${formatPathNumber(height)}`,
      `A ${formatPathNumber(width * 0.42)} ${formatPathNumber(
        height * 0.43,
      )} 0 1 1 ${formatPathNumber(startX)} 0`,
      'Z',
    ].join(' ');
  }

  const point = (x: number, y: number) =>
    String(formatPathNumber(width * x)) +
    ' ' +
    String(formatPathNumber(height * y));
  if (preset === 'cloud') {
    // 云朵由 OOXML 预设中的 11 段椭圆弧组成，不能使用近似贝塞尔轮廓代替。
    const scaleX = width / 43200;
    const scaleY = height / 43200;
    const arcs = [
      [6753, 9190, -11429249, 7426832],
      [5333, 7267, -8646143, 5396714],
      [4365, 5945, -8748475, 5983381],
      [4857, 6595, -7859164, 7034504],
      [5333, 7273, -4722533, 6541615],
      [6775, 9220, -2776035, 7816140],
      [5785, 7867, 37501, 6842000],
      [6752, 9215, 1347096, 6910353],
      [7720, 10543, 3974558, 4542661],
      [4360, 5918, -16496525, 8804134],
      [4345, 5945, -14809710, 9151131],
    ] as const;
    let currentX = 3900;
    let currentY = 14370;
    const commands = [
      'M ' +
        formatPathNumber(currentX * scaleX) +
        ' ' +
        formatPathNumber(currentY * scaleY),
    ];

    arcs.forEach(([radiusX, radiusY, startAngle, sweepAngle]) => {
      const startRadians = (startAngle / 60000 / 180) * Math.PI;
      const endRadians = ((startAngle + sweepAngle) / 60000 / 180) * Math.PI;
      const startParameter = Math.atan2(
        radiusX * Math.sin(startRadians),
        radiusY * Math.cos(startRadians),
      );
      const endParameter = Math.atan2(
        radiusX * Math.sin(endRadians),
        radiusY * Math.cos(endRadians),
      );
      const centerX = currentX - radiusX * Math.cos(startParameter);
      const centerY = currentY - radiusY * Math.sin(startParameter);
      currentX = centerX + radiusX * Math.cos(endParameter);
      currentY = centerY + radiusY * Math.sin(endParameter);
      commands.push(
        [
          'A',
          formatPathNumber(radiusX * scaleX),
          formatPathNumber(radiusY * scaleY),
          0,
          Math.abs(sweepAngle) > 10800000 ? 1 : 0,
          sweepAngle > 0 ? 1 : 0,
          formatPathNumber(currentX * scaleX),
          formatPathNumber(currentY * scaleY),
        ].join(' '),
      );
    });

    commands.push('Z');
    return commands.join(' ');
  }
  if (preset === 'heart') {
    return [
      `M ${point(0.5, 1)}`,
      `C ${point(0.43, 0.91)} ${point(0, 0.64)} ${point(0, 0.31)}`,
      `C ${point(0, 0.06)} ${point(0.29, -0.04)} ${point(0.5, 0.21)}`,
      `C ${point(0.71, -0.04)} ${point(1, 0.06)} ${point(1, 0.31)}`,
      `C ${point(1, 0.64)} ${point(0.57, 0.91)} ${point(0.5, 1)}`,
      'Z',
    ].join(' ');
  }

  return undefined;
}

function parseWpgShapeItem(
  shapeNode: Element,
  index: number,
  context: ParseContext,
  matrix: DrawingMatrix,
  readBlocks: DocxDrawingContentReader,
): DocxShapeItem | undefined {
  const spPr = childByLocalName(shapeNode, 'spPr');
  const size = readDrawingItemSize(spPr, matrix);
  const kind = readDrawingShapeKind(spPr);
  const isLine = kind === 'line';
  if (!isLine && (!size.width || !size.height)) return undefined;
  if (isLine && !size.width && !size.height) return undefined;

  const id = `wpg-item-${context.shapeIndex + 1}-${index + 1}`;
  const fillColor = parseDrawingFillColor(spPr, context.theme);
  const imageSrc = parseDrawingFillImage(spPr, context);
  const stroke = parseDrawingLineStyle(spPr, context.theme);
  const textBehavior = readDrawingTextBehavior(shapeNode);
  const blocks = retainTextBoxLayoutBlocks(
    parseVmlTextBoxParagraphs(shapeNode, context, id, readBlocks),
  );
  const pathLayers = convertHorizontalScrollPathLayers(
    spPr,
    size.width,
    size.height,
    fillColor,
    stroke,
  );
  const path =
    (pathLayers
      ? undefined
      : convertDrawingPresetGeometry(spPr, size.width, size.height)) ??
    (kind === 'path'
      ? convertDrawingCustomGeometry(spPr, size.width, size.height)
      : isLine
      ? `M 0 0 L ${size.width || 0} ${size.height || 0}`
      : undefined);

  return {
    id,
    kind,
    ...size,
    ...readDrawingItemTransform(spPr, matrix),
    height: isLine && size.height === 0 ? 1 : size.height,
    ...readDrawingBodyPadding(shapeNode),
    path,
    pathLayers,
    viewBox:
      path || pathLayers?.length
        ? `0 0 ${Math.max(1, size.width)} ${Math.max(1, size.height)}`
        : undefined,
    fillColor,
    imageSrc,
    ...stroke,
    borderRadius:
      kind === 'ellipse' ? '50%' : readDrawingShapeBorderRadius(spPr, size),
    textVerticalAlign: readDrawingTextAnchor(shapeNode),
    fitShapeToText: textBehavior.fitShapeToText || undefined,
    clipVerticalOverflow: textBehavior.clipVerticalOverflow || undefined,
    noWrap: textBehavior.noWrap || undefined,
    hyperlink: parseDocxDrawingHyperlink(shapeNode, context.documentRels),
    blocks: blocks.length ? blocks : undefined,
    paragraphs: blocks.filter(
      (block): block is DocxParagraphBlock => block.type === 'paragraph',
    ),
  };
}

function parseWpgPictureItem(
  pictureNode: Element,
  index: number,
  context: ParseContext,
  matrix: DrawingMatrix,
): DocxShapeItem | undefined {
  const spPr = childByLocalName(pictureNode, 'spPr');
  const size = readDrawingItemSize(spPr, matrix);
  if (!size.width || !size.height) return undefined;

  const blip = childByLocalName(
    childByLocalName(pictureNode, 'blipFill'),
    'blip',
  );
  const embed = attr(blip, 'r:embed') ?? attr(blip, 'embed');
  const target = embed ? context.documentRels[embed]?.target : undefined;
  const imageSrc = resolveMediaRef(target, context.packageState);
  if (!imageSrc) return undefined;

  const kind = readDrawingShapeKind(spPr);
  const path =
    convertDrawingPresetGeometry(spPr, size.width, size.height) ??
    (kind === 'path'
      ? convertDrawingCustomGeometry(spPr, size.width, size.height)
      : undefined);

  return {
    id: `wpg-picture-item-${context.shapeIndex + 1}-${index + 1}`,
    kind,
    ...size,
    ...readDrawingItemTransform(spPr, matrix),
    path,
    viewBox: path
      ? `0 0 ${Math.max(1, size.width)} ${Math.max(1, size.height)}`
      : undefined,
    imageSrc,
    ...parseDrawingLineStyle(spPr, context.theme),
    borderRadius:
      kind === 'ellipse' ? '50%' : readDrawingShapeBorderRadius(spPr, size),
    hyperlink: parseDocxDrawingHyperlink(pictureNode, context.documentRels),
  };
}

function parseWpgShape(
  node: Element,
  context: ParseContext,
  readBlocks: DocxDrawingContentReader,
): DocxShape | undefined {
  const group = descendantByLocalName(node, 'wgp');
  const extent = descendantByLocalName(node, 'extent');
  const width = Math.round(emuToPx(Number(attr(extent, 'cx') ?? 0)));
  const height = Math.round(emuToPx(Number(attr(extent, 'cy') ?? 0)));
  if (!width || !height) return undefined;

  let items: DocxShapeItem[];
  if (group) {
    const drawables = collectWpgDrawables(
      group,
      readWpgRootMatrix(group, width, height),
    );
    items = drawables
      .map(({ node: child, matrix }, index) => {
        if (matchesLocalName(child, 'wsp')) {
          return parseWpgShapeItem(child, index, context, matrix, readBlocks);
        }
        if (matchesLocalName(child, 'pic')) {
          return parseWpgPictureItem(child, index, context, matrix);
        }
        return undefined;
      })
      .filter((item): item is DocxShapeItem => Boolean(item));
  } else {
    // 无 wgp 包装的独立 wsp，作为整个锚点尺寸的单元素形状处理
    const graphicData = descendantByLocalName(node, 'graphicData');
    const standaloneWsp = graphicData
      ? childByLocalName(graphicData, 'wsp')
      : undefined;
    if (!standaloneWsp) return undefined;
    const emuMatrix = createDrawingScaleMatrix(emuToPx(1), emuToPx(1));
    const item = parseWpgShapeItem(
      standaloneWsp,
      0,
      context,
      emuMatrix,
      readBlocks,
    );
    if (!item) return undefined;
    // 独立 wsp 的锚点已经提供整体位置，内部 xfrm 只描述该形状自身尺寸，不能再次偏移。
    items = [{ ...item, left: 0, top: 0 }];
  }

  if (!items.length) return undefined;
  context.shapeIndex += 1;
  const anchorPosition = readDrawingAnchorPosition(
    node,
    context.documentGridLineHeight,
  );
  return {
    id: `docx-shape-${context.shapeIndex}`,
    width,
    height,
    position: anchorPosition,
    items,
  };
}

function parseAlternateContentShape(
  node: Element,
  context: ParseContext,
  readBlocks: DocxDrawingContentReader,
): DocxShape | undefined {
  const choice = childByLocalName(node, 'Choice');
  const choiceDrawing = descendantByLocalName(choice, 'drawing');
  const choiceShape = choiceDrawing
    ? parseWpgShape(choiceDrawing, context, readBlocks)
    : undefined;
  const fallback = childByLocalName(node, 'Fallback');
  const fallbackPict = descendantByLocalName(fallback, 'pict');

  if (choiceShape) {
    const fallbackPosition = readVmlShapeContainerPosition(fallbackPict);
    // AlternateContent 的 Choice 是当前解析器支持的主表示，Fallback 仅补齐缺失定位。
    const mergedPosition = mergeDocxPosition(
      fallbackPosition,
      choiceShape.position,
    );
    return {
      ...choiceShape,
      position: mergedPosition,
    };
  }

  return fallbackPict
    ? parseVmlShape(fallbackPict, context, readBlocks)
    : undefined;
}

/** 合并 `mergeDocxPosition` 接收的多份数据。 */
function mergeDocxPosition(
  base: DocxPosition | undefined,
  override: DocxPosition | undefined,
): DocxPosition | undefined {
  if (!base) return override;
  if (!override) return base;
  return {
    ...base,
    ...override,
    zIndex: override.zIndex ?? base.zIndex,
    behindDoc: override.behindDoc ?? base.behindDoc,
  };
}

function parseVmlCoordSize(
  node: Element,
  renderedWidth: number,
  renderedHeight: number,
) {
  const coordsize = attr(node, 'coordsize');
  const [coordWidth, coordHeight] = (coordsize ?? '')
    .split(',')
    .map((value) => Number(value.trim()));
  return {
    x:
      Number.isFinite(coordWidth) && coordWidth > 0
        ? renderedWidth / coordWidth
        : undefined,
    y:
      Number.isFinite(coordHeight) && coordHeight > 0
        ? renderedHeight / coordHeight
        : undefined,
  };
}

function readVmlCoordOrigin(node: Element | null | undefined) {
  const [x, y] = (attr(node, 'coordorigin') ?? '')
    .split(',')
    .map((value) => Number(value.trim()));
  return {
    x: Number.isFinite(x) ? x : 0,
    y: Number.isFinite(y) ? y : 0,
  };
}

function readVmlCoordSize(node: Element) {
  const [width, height] = (attr(node, 'coordsize') ?? '')
    .split(',')
    .map((value) => Number(value.trim()));
  return {
    width: Number.isFinite(width) && width > 0 ? width : undefined,
    height: Number.isFinite(height) && height > 0 ? height : undefined,
  };
}

function parseVmlShapeSize(
  node: Element,
  scale?: {
    /** 水平方向的坐标或缩放参数。 */
    x?: number;
    /** 垂直方向的坐标或缩放参数。 */
    y?: number;
  },
) {
  const style = attr(node, 'style');
  return {
    left: readCssSize(style, 'left', scale?.x) ?? 0,
    top: readCssSize(style, 'top', scale?.y) ?? 0,
    width: readCssSize(style, 'width', scale?.x) ?? 0,
    height: readCssSize(style, 'height', scale?.y) ?? 0,
  };
}

function readVmlShapePosition(node: Element | null | undefined) {
  const style = attr(node, 'style');
  const left = readCssPosition(style, 'left');
  const top = readCssPosition(style, 'top');
  const zIndex = Number(readCssDeclaration(style, 'z-index'));
  const rotation = readCssDeclaration(style, 'rotation');
  if (left === undefined && top === undefined) return undefined;
  const isBehindDoc = Number.isFinite(zIndex) && zIndex < 0;
  return {
    left: Math.round(left ?? 0),
    top: Math.round(top ?? 0),
    relativeFromH: readCssDeclaration(
      style,
      'mso-position-horizontal-relative',
    ) as DocxPosition['relativeFromH'],
    relativeFromV: readCssDeclaration(
      style,
      'mso-position-vertical-relative',
    ) as DocxPosition['relativeFromV'],
    zIndex: Number.isFinite(zIndex) && zIndex >= 0 ? zIndex : undefined,
    behindDoc: isBehindDoc || undefined,
    rotation: rotation ? Number(rotation) : undefined,
    flipH:
      readCssDeclaration(style, 'flip') === 'x' ||
      readCssDeclaration(style, 'flip') === 'xy' ||
      undefined,
    flipV:
      readCssDeclaration(style, 'flip') === 'y' ||
      readCssDeclaration(style, 'flip') === 'xy' ||
      undefined,
  };
}

function readVmlShapeContainerPosition(node: Element | null | undefined) {
  if (!node) return undefined;
  const group = matchesLocalName(node, 'group')
    ? node
    : descendantByLocalName(node, 'group');
  if (group) return readVmlShapePosition(group);
  const shape = Array.from(node.children).find(
    (child) =>
      matchesLocalName(child, 'shape') ||
      matchesLocalName(child, 'rect') ||
      matchesLocalName(child, 'roundrect'),
  );
  return readVmlShapePosition(shape ?? node);
}

function vmlOnOff(value: string | undefined, fallback = true) {
  if (value === undefined) return fallback;
  return value !== 'f' && value !== 'false' && value !== '0' && value !== 'off';
}

function parseVmlStroke(shapeNode: Element) {
  const stroke = childByLocalName(shapeNode, 'stroke');
  const stroked = attr(shapeNode, 'stroked');
  const strokeOn = attr(stroke, 'on');

  if (!vmlOnOff(stroked, true) || !vmlOnOff(strokeOn, true)) {
    return {};
  }

  const rawColor = attr(stroke, 'color') ?? attr(shapeNode, 'strokecolor');
  const color = normalizeCssColor(rawColor) ?? '#000';
  const width = vmlUnitToPx(attr(stroke, 'weight')) ?? 1;
  const dashstyle = attr(stroke, 'dashstyle');
  const strokeDasharray = dashstyle
    ?.split(/\s+/)
    .map((item) => Number(item))
    .filter((item) => Number.isFinite(item) && item > 0)
    .map((item) => item * width)
    .join(' ');

  const result = {
    border: `${width}px solid ${color}`,
    strokeColor: color,
    strokeWidth: width,
    strokeDasharray: strokeDasharray || undefined,
  };

  return result;
}

function parseVmlFillColor(shapeNode: Element) {
  const fill = childByLocalName(shapeNode, 'fill');
  if (
    !vmlOnOff(attr(shapeNode, 'filled'), true) ||
    !vmlOnOff(attr(fill, 'on'), true)
  ) {
    return undefined;
  }
  return normalizeCssColor(attr(fill, 'color') ?? attr(shapeNode, 'fillcolor'));
}

function readVmlTextAnchor(
  shapeNode: Element,
): DocxShapeItem['textVerticalAlign'] {
  const anchor = readCssDeclaration(attr(shapeNode, 'style'), 'v-text-anchor');
  if (anchor === 'middle') return 'middle';
  if (anchor === 'bottom') return 'bottom';
  return 'top';
}

function readVmlTextBehavior(shapeNode: Element) {
  const shapeStyle = attr(shapeNode, 'style');
  const textboxNode = descendantByLocalName(shapeNode, 'textbox');
  const textboxStyle = attr(textboxNode, 'style');
  const fitShapeToText =
    readCssDeclaration(textboxStyle, 'mso-fit-shape-to-text') === 't';

  return {
    fitShapeToText,
    noWrap:
      fitShapeToText ||
      readCssDeclaration(shapeStyle, 'mso-wrap-style') === 'none',
  };
}

function convertVmlPathToSvgPath(
  path: string | undefined,
  width: number,
  height: number,
  node: Element,
) {
  if (!path) return undefined;
  const coordSize = readVmlCoordSize(node);
  const scaleX = coordSize.width ? width / coordSize.width : 1;
  const scaleY = coordSize.height ? height / coordSize.height : 1;
  const tokens = path.match(/[a-z]|-?\d+(?:\.\d+)?/gi) ?? [];
  const commands: string[] = [];
  let index = 0;
  let command = '';

  const readPoint = () => {
    const x = Number(tokens[index++]);
    const y = Number(tokens[index++]);
    if (!Number.isFinite(x) || !Number.isFinite(y)) return undefined;
    return `${x * scaleX} ${y * scaleY}`;
  };

  while (index < tokens.length) {
    const token = tokens[index++];
    if (/^[a-z]$/i.test(token)) {
      command = token.toLowerCase();
    } else {
      index -= 1;
    }

    if (command === 'm') {
      const point = readPoint();
      if (point) commands.push(`M ${point}`);
    } else if (command === 'l') {
      const point = readPoint();
      if (point) commands.push(`L ${point}`);
    } else if (command === 'c') {
      const points = [readPoint(), readPoint(), readPoint()];
      if (points.every(Boolean)) commands.push(`C ${points.join(' ')}`);
    } else if (command === 'x') {
      commands.push('Z');
    } else if (command === 'e') {
      break;
    } else {
      break;
    }
  }

  return commands.length ? commands.join(' ') : undefined;
}

function parseVmlTextBoxParagraphs(
  shapeNode: Element,
  context: ParseContext,
  id: string,
  readBlocks: DocxDrawingContentReader,
) {
  const textBox = descendantByLocalName(shapeNode, 'txbxContent');
  return readBlocks(textBox, id, context, { insideShape: true });
}

/** 保留文本框内部的可见内容；末尾空段落只表示编辑终点，不额外占用版面。 */
function retainTextBoxLayoutBlocks(blocks: DocxBlock[]) {
  const hasVisibleContent = blocks.some(
    (block) =>
      block.type !== 'paragraph' || Boolean(block.text || block.inlines.length),
  );
  if (!hasVisibleContent) return [];

  let finalBlockIndex = blocks.length - 1;
  while (finalBlockIndex >= 0) {
    const finalBlock = blocks[finalBlockIndex];
    if (
      finalBlock.type !== 'paragraph' ||
      finalBlock.text ||
      finalBlock.inlines.length
    ) {
      break;
    }
    finalBlockIndex -= 1;
  }
  return blocks.slice(0, finalBlockIndex + 1);
}

function hasVmlTextBox(shapeNode: Element) {
  return Boolean(descendantByLocalName(shapeNode, 'txbxContent'));
}

function readVmlShapeBorderRadius(
  shapeNode: Element,
  size: { width: number; height: number },
) {
  if (!matchesLocalName(shapeNode, 'roundrect')) return undefined;
  const rawArcSize = attr(shapeNode, 'arcsize')?.trim();
  if (!rawArcSize) return 8;
  const numericArcSize = Number.parseFloat(rawArcSize);
  if (!Number.isFinite(numericArcSize)) return 8;
  const ratio = rawArcSize.endsWith('%')
    ? numericArcSize / 100
    : numericArcSize;
  return Math.min(
    Math.min(size.width, size.height) / 2,
    Math.max(0, Math.min(size.width, size.height) * ratio),
  );
}

function parseVmlShapeItem(
  shapeNode: Element,
  index: number,
  context: ParseContext,
  readBlocks: DocxDrawingContentReader,
  scale?: {
    /** 水平方向的坐标或缩放参数。 */
    x?: number;
    /** 垂直方向的坐标或缩放参数。 */
    y?: number;
  },
  origin?: {
    /** 水平方向的坐标或缩放参数。 */
    x: number;
    /** 垂直方向的坐标或缩放参数。 */
    y: number;
  },
): DocxShapeItem | undefined {
  const size = parseVmlShapeSize(shapeNode, scale);
  size.left -= (origin?.x ?? 0) * (scale?.x ?? 0);
  size.top -= (origin?.y ?? 0) * (scale?.y ?? 0);
  if (!size.width || !size.height) return undefined;

  const isEllipse =
    matchesLocalName(shapeNode, 'shape') &&
    ((attr(shapeNode, 'o:spt') ?? attr(shapeNode, 'spt')) === '3' ||
      (attr(shapeNode, 'type') ?? '').includes('_x0000_t3'));
  const fillColor = parseVmlFillColor(shapeNode);
  const stroke = parseVmlStroke(shapeNode);
  const path = convertVmlPathToSvgPath(
    attr(shapeNode, 'path'),
    size.width,
    size.height,
    shapeNode,
  );
  const id = `vml-item-${context.shapeIndex + 1}-${index + 1}`;
  const blocks = retainTextBoxLayoutBlocks(
    parseVmlTextBoxParagraphs(shapeNode, context, id, readBlocks),
  );

  const textBehavior = readVmlTextBehavior(shapeNode);

  return {
    id,
    kind: isEllipse ? 'ellipse' : 'rect',
    ...size,
    path,
    viewBox: path ? `0 0 ${size.width} ${size.height}` : undefined,
    fillColor,
    ...stroke,
    borderRadius: isEllipse ? '50%' : readVmlShapeBorderRadius(shapeNode, size),
    textVerticalAlign: readVmlTextAnchor(shapeNode),
    fitShapeToText: textBehavior.fitShapeToText || undefined,
    noWrap: textBehavior.noWrap || undefined,
    hyperlink: parseDocxDrawingHyperlink(shapeNode, context.documentRels),
    blocks: blocks.length ? blocks : undefined,
    paragraphs: blocks.filter(
      (block): block is DocxParagraphBlock => block.type === 'paragraph',
    ),
  };
}

function parseVmlShape(
  node: Element,
  context: ParseContext,
  readBlocks: DocxDrawingContentReader,
): DocxShape | undefined {
  const group = matchesLocalName(node, 'group')
    ? node
    : descendantByLocalName(node, 'group');
  const shapeRoot = group ?? node;
  const rootSize = parseVmlShapeSize(shapeRoot);
  const scale = parseVmlCoordSize(shapeRoot, rootSize.width, rootSize.height);
  const origin = readVmlCoordOrigin(shapeRoot);

  // 如果 shapeRoot 是 pict，查找其中的 shape 子元素
  const rawItems = Array.from(shapeRoot.children).filter(
    (child) =>
      (matchesLocalName(child, 'shape') ||
        matchesLocalName(child, 'rect') ||
        matchesLocalName(child, 'roundrect')) &&
      (child.hasAttribute('fillcolor') ||
        child.hasAttribute('strokecolor') ||
        attr(child, 'filled') !== 'f' ||
        attr(child, 'stroked') !== 'f' ||
        hasVmlTextBox(child)),
  );

  const position = group
    ? readVmlShapePosition(shapeRoot)
    : readVmlShapePosition(rawItems[0] ?? shapeRoot);
  const items = rawItems
    .map((child, index) =>
      parseVmlShapeItem(child, index, context, readBlocks, scale, origin),
    )
    .filter((item): item is DocxShapeItem => Boolean(item));

  if (!items.length) return undefined;

  const maxRight = Math.max(...items.map((item) => item.left + item.width));
  const maxBottom = Math.max(...items.map((item) => item.top + item.height));
  context.shapeIndex += 1;

  return {
    id: `docx-shape-${context.shapeIndex}`,
    width: rootSize.width || maxRight,
    height: rootSize.height || maxBottom,
    position,
    items,
  };
}

/** 创建可递归读取文本框内容的 DOCX 绘图解析器。 */
export function createDocxDrawingParser(
  readBlocks: DocxDrawingContentReader,
): DocxDrawingParser {
  return {
    parseRunChild(node, context) {
      if (matchesLocalName(node, 'drawing')) {
        const shape = parseWpgShape(node, context, readBlocks);
        if (shape) return { type: 'shape', shape };

        const webExtensionChart = parseWpsWebExtensionChart(node, context);
        if (webExtensionChart) {
          return { type: 'chart', chart: webExtensionChart };
        }

        const chart = parseChartElement(node, context);
        if (chart) return { type: 'chart', chart };

        const image = isDirectDrawingPicture(node)
          ? parseDrawingImage(node, context)
          : undefined;
        return image ? { type: 'image', image } : undefined;
      }

      if (
        !matchesLocalName(node, 'pict') &&
        !matchesLocalName(node, 'alternateContent')
      ) {
        return undefined;
      }

      if (matchesLocalName(node, 'alternateContent')) {
        const drawing = descendantByLocalName(node, 'drawing');
        const image = drawing
          ? parseAlternateContentImage(drawing, context)
          : undefined;
        if (image) {
          const fallbackPict = descendantByLocalName(
            childByLocalName(node, 'Fallback'),
            'pict',
          );
          const position = readVmlShapeContainerPosition(fallbackPict);
          return {
            type: 'image',
            image: position ? { ...image, position } : image,
          };
        }
      }

      const shape = matchesLocalName(node, 'pict')
        ? parseVmlShape(node, context, readBlocks)
        : parseAlternateContentShape(node, context, readBlocks);
      return shape ? { type: 'shape', shape } : undefined;
    },
  };
}
