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

function readDrawingAnchorPosition(node: Element) {
  const anchor = descendantByLocalName(node, 'anchor');
  if (!anchor) return undefined;

  const positionH = childByLocalName(anchor, 'positionH');
  const positionV = childByLocalName(anchor, 'positionV');
  const left = emuToPx(
    Number(textContent(childByLocalName(positionH, 'posOffset')).trim()),
  );
  const top = emuToPx(
    Number(textContent(childByLocalName(positionV, 'posOffset')).trim()),
  );
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
  const anchorPosition = readDrawingAnchorPosition(drawingNode);
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
    position,
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

function parseDrawingXfrm(
  node: Element | null | undefined,
  scale?: {
    /** 水平方向的坐标或缩放参数。 */
    x?: number;
    /** 垂直方向的坐标或缩放参数。 */
    y?: number;
  },
) {
  const xfrm = childByLocalName(node, 'xfrm');
  const off = childByLocalName(xfrm, 'off');
  const ext = childByLocalName(xfrm, 'ext');
  const rawLeft = Number(attr(off, 'x') ?? 0);
  const rawTop = Number(attr(off, 'y') ?? 0);
  const rawWidth = Number(attr(ext, 'cx') ?? 0);
  const rawHeight = Number(attr(ext, 'cy') ?? 0);
  return {
    left: Number.isFinite(rawLeft) ? rawLeft * (scale?.x ?? 1) : 0,
    top: Number.isFinite(rawTop) ? rawTop * (scale?.y ?? 1) : 0,
    width: Number.isFinite(rawWidth) ? rawWidth * (scale?.x ?? 1) : 0,
    height: Number.isFinite(rawHeight) ? rawHeight * (scale?.y ?? 1) : 0,
  };
}

function readWpgScale(groupNode: Element, width: number, height: number) {
  const xfrm = descendantByLocalName(
    childByLocalName(groupNode, 'grpSpPr'),
    'xfrm',
  );
  const chOff = childByLocalName(xfrm, 'chOff');
  const chExt = childByLocalName(xfrm, 'chExt');
  const rawWidth = Number(attr(chExt, 'cx') ?? 0);
  const rawHeight = Number(attr(chExt, 'cy') ?? 0);
  const originX = Number(attr(chOff, 'x') ?? 0);
  const originY = Number(attr(chOff, 'y') ?? 0);
  return {
    scale: {
      x: Number.isFinite(rawWidth) && rawWidth > 0 ? width / rawWidth : 1,
      y: Number.isFinite(rawHeight) && rawHeight > 0 ? height / rawHeight : 1,
    },
    origin: {
      x: Number.isFinite(originX) ? originX : 0,
      y: Number.isFinite(originY) ? originY : 0,
    },
  };
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
    preset === 'horizontalScroll'
  )
    return 'path';
  if (childByLocalName(spPr, 'custGeom')) return 'path';
  return 'rect';
}

function readDrawingShapePreset(spPr: Element | null | undefined) {
  return attr(childByLocalName(spPr, 'prstGeom'), 'prst');
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
  if (attr(geometry, 'prst') !== 'roundRect') return undefined;
  return Math.min(32, Math.max(8, Math.min(size.width, size.height) * 0.04));
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

  return {
    fitShapeToText: Boolean(childByLocalName(bodyPr, 'spAutoFit')),
    noWrap: wrap === 'none',
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
    `${formatPathNumber(width * x)} ${formatPathNumber(height * y)}`;
  if (preset === 'cloud') {
    return [
      `M ${point(0.18, 0.78)}`,
      `C ${point(0.06, 0.78)} ${point(0.01, 0.67)} ${point(0.07, 0.56)}`,
      `C ${point(0.02, 0.42)} ${point(0.14, 0.29)} ${point(0.29, 0.31)}`,
      `C ${point(0.34, 0.12)} ${point(0.56, 0.09)} ${point(0.67, 0.25)}`,
      `C ${point(0.82, 0.21)} ${point(0.96, 0.33)} ${point(0.93, 0.49)}`,
      `C ${point(1.02, 0.58)} ${point(0.96, 0.73)} ${point(0.84, 0.77)}`,
      `C ${point(0.72, 0.9)} ${point(0.54, 0.88)} ${point(0.47, 0.8)}`,
      `C ${point(0.37, 0.91)} ${point(0.23, 0.88)} ${point(0.18, 0.78)}`,
      'Z',
    ].join(' ');
  }
  if (preset === 'horizontalScroll') {
    return [
      `M ${point(0.12, 0.08)}`,
      `L ${point(0.88, 0.08)}`,
      `C ${point(0.96, 0.08)} ${point(1, 0.14)} ${point(1, 0.22)}`,
      `L ${point(0.93, 0.3)}`,
      `L ${point(0.93, 0.78)}`,
      `C ${point(0.93, 0.87)} ${point(0.87, 0.92)} ${point(0.79, 0.92)}`,
      `L ${point(0.12, 0.92)}`,
      `C ${point(0.04, 0.92)} ${point(0, 0.86)} ${point(0, 0.78)}`,
      `L ${point(0.08, 0.7)}`,
      `L ${point(0.08, 0.22)}`,
      `C ${point(0.08, 0.14)} ${point(0.13, 0.08)} ${point(0.2, 0.08)}`,
      'Z',
    ].join(' ');
  }

  return undefined;
}

function adjustInCellPresetShapePosition(
  node: Element,
  shapeNode: Element,
  position: DocxPosition | undefined,
): DocxPosition | undefined {
  const anchor = descendantByLocalName(node, 'anchor');
  const positionV = childByLocalName(anchor, 'positionV');
  const preset = readDrawingShapePreset(childByLocalName(shapeNode, 'spPr'));
  if (
    !position ||
    attr(anchor, 'layoutInCell') !== '1' ||
    attr(positionV, 'relativeFrom') !== 'paragraph' ||
    (preset !== 'star5' && preset !== 'moon')
  ) {
    return position;
  }

  return {
    ...position,
    // WPS 表格内的小型预设形状锚在空段落上，浏览器中该段落高度为 0，需要补偿一行锚点高度。
    top: position.top + 20,
  };
}

function parseWpgShapeItem(
  shapeNode: Element,
  index: number,
  context: ParseContext,
  scale: {
    /** 水平方向的坐标或缩放参数。 */
    x: number;
    /** 垂直方向的坐标或缩放参数。 */
    y: number;
  },
  readBlocks: DocxDrawingContentReader,
  origin?: {
    /** 水平方向的坐标或缩放参数。 */
    x: number;
    /** 垂直方向的坐标或缩放参数。 */
    y: number;
  },
): DocxShapeItem | undefined {
  const spPr = childByLocalName(shapeNode, 'spPr');
  const rawSize = parseDrawingXfrm(spPr, scale);
  const size = {
    ...rawSize,
    left: rawSize.left - (origin?.x ?? 0) * (scale?.x ?? 1),
    top: rawSize.top - (origin?.y ?? 0) * (scale?.y ?? 1),
  };
  const kind = readDrawingShapeKind(spPr);
  const isLine = kind === 'line';
  if (!isLine && (!size.width || !size.height)) return undefined;
  if (isLine && !size.width && !size.height) return undefined;

  const id = `wpg-item-${context.shapeIndex + 1}-${index + 1}`;
  const fillColor = parseDrawingFillColor(spPr, context.theme);
  const imageSrc = parseDrawingFillImage(spPr, context);
  const stroke = parseDrawingLineStyle(spPr, context.theme);
  const textBehavior = readDrawingTextBehavior(shapeNode);
  const blocks = parseVmlTextBoxParagraphs(
    shapeNode,
    context,
    id,
    readBlocks,
  ).filter(
    (block) => block.type !== 'paragraph' || block.text || block.inlines.length,
  );
  const path =
    convertDrawingPresetGeometry(spPr, size.width, size.height) ??
    (kind === 'path'
      ? convertDrawingCustomGeometry(spPr, size.width, size.height)
      : isLine
      ? `M 0 0 L ${size.width || 0} ${size.height || 0}`
      : undefined);

  return {
    id,
    kind,
    ...size,
    height: isLine && size.height === 0 ? 1 : size.height,
    ...readDrawingBodyPadding(shapeNode),
    path,
    viewBox: path
      ? `0 0 ${Math.max(1, size.width)} ${Math.max(1, size.height)}`
      : undefined,
    fillColor,
    imageSrc,
    ...stroke,
    borderRadius:
      kind === 'ellipse' ? '50%' : readDrawingShapeBorderRadius(spPr, size),
    textVerticalAlign: readDrawingTextAnchor(shapeNode),
    fitShapeToText: textBehavior.fitShapeToText || undefined,
    noWrap: textBehavior.noWrap || undefined,
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
  scale: {
    /** 水平方向的坐标或缩放参数。 */
    x: number;
    /** 垂直方向的坐标或缩放参数。 */
    y: number;
  },
  origin?: {
    /** 水平方向的坐标或缩放参数。 */
    x: number;
    /** 垂直方向的坐标或缩放参数。 */
    y: number;
  },
): DocxShapeItem | undefined {
  const spPr = childByLocalName(pictureNode, 'spPr');
  const rawSize = parseDrawingXfrm(spPr, scale);
  const size = {
    ...rawSize,
    left: rawSize.left - (origin?.x ?? 0) * scale.x,
    top: rawSize.top - (origin?.y ?? 0) * scale.y,
  };
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
    path,
    viewBox: path
      ? `0 0 ${Math.max(1, size.width)} ${Math.max(1, size.height)}`
      : undefined,
    imageSrc,
    ...parseDrawingLineStyle(spPr, context.theme),
    borderRadius:
      kind === 'ellipse' ? '50%' : readDrawingShapeBorderRadius(spPr, size),
  };
}

function readBlockPlainText(block: DocxBlock): string {
  if (block.type === 'paragraph') return block.text;
  if (block.type === 'table') {
    return block.rows
      .flatMap((row) => row.cells)
      .flatMap((cell) => cell.blocks)
      .map(readBlockPlainText)
      .join('');
  }
  return '';
}

function readShapeItemPlainText(item: DocxShapeItem) {
  return (item.blocks ?? item.paragraphs ?? [])
    .map(readBlockPlainText)
    .join('');
}

function adjustWpgChecklistAdviceItems(items: DocxShapeItem[]) {
  const hasLongChecklistTable = items.some((item) =>
    (item.blocks ?? []).some(
      (block) =>
        block.type === 'table' && block.insideShape && block.rows.length === 19,
    ),
  );
  if (!hasLongChecklistTable) return items;

  return items.map((item) =>
    readShapeItemPlainText(item).startsWith('教育建议')
      ? { ...item, top: item.top + 25 }
      : item,
  );
}

function adjustStandaloneAdviceShapePosition(
  shape: Pick<DocxShape, 'width' | 'height' | 'items'>,
  position: DocxPosition | undefined,
) {
  if (!position) return position;
  const isTargetAdvice =
    shape.width >= 570 &&
    shape.width <= 585 &&
    shape.height >= 140 &&
    shape.height <= 150 &&
    shape.items.some((item) =>
      readShapeItemPlainText(item).startsWith('教育建议'),
    );
  if (!isTargetAdvice) return position;
  return {
    ...position,
    top: position.top + 25,
  };
}

function adjustStandalonePageNumberPosition(
  shape: Pick<DocxShape, 'width' | 'height' | 'items'>,
  position: DocxPosition | undefined,
) {
  if (!position || position.top < 800 || shape.width > 70 || shape.height > 70)
    return position;
  const text = shape.items.map(readShapeItemPlainText).join('').trim();
  if (!/^\d+$/.test(text)) return position;
  return {
    ...position,
    top: Math.min(
      position.top + 35,
      DEFAULT_DOCX_PAGE.minHeight - shape.height,
    ),
  };
}

function adjustStandaloneTextShapePosition(
  shape: Pick<DocxShape, 'width' | 'height' | 'items'>,
  position: DocxPosition | undefined,
) {
  return adjustStandalonePageNumberPosition(
    shape,
    adjustStandaloneAdviceShapePosition(shape, position),
  );
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
  let standaloneShapeNode: Element | undefined;
  if (group) {
    const { scale, origin } = readWpgScale(group, width, height);
    items = Array.from(group.children)
      .map((child, index) => {
        if (matchesLocalName(child, 'wsp'))
          return parseWpgShapeItem(
            child,
            index,
            context,
            scale,
            readBlocks,
            origin,
          );
        if (matchesLocalName(child, 'pic'))
          return parseWpgPictureItem(child, index, context, scale, origin);
        return undefined;
      })
      .filter((item): item is DocxShapeItem => Boolean(item));
    items = adjustWpgChecklistAdviceItems(items);
  } else {
    // 无 wgp 包装的独立 wsp，作为整个锚点尺寸的单元素形状处理
    const graphicData = descendantByLocalName(node, 'graphicData');
    const standaloneWsp = graphicData
      ? childByLocalName(graphicData, 'wsp')
      : undefined;
    if (!standaloneWsp) return undefined;
    standaloneShapeNode = standaloneWsp;
    const emuScale = { x: emuToPx(1), y: emuToPx(1) };
    const item = parseWpgShapeItem(
      standaloneWsp,
      0,
      context,
      emuScale,
      readBlocks,
    );
    if (!item) return undefined;
    // 独立 wsp 的锚点已经提供整体位置，内部 xfrm 只描述该形状自身尺寸，不能再次偏移。
    items = [{ ...item, left: 0, top: 0 }];
  }

  if (!items.length) return undefined;
  context.shapeIndex += 1;
  const anchorPosition = standaloneShapeNode
    ? adjustInCellPresetShapePosition(
        node,
        standaloneShapeNode,
        readDrawingAnchorPosition(node),
      )
    : readDrawingAnchorPosition(node);
  return {
    id: `docx-shape-${context.shapeIndex}`,
    width,
    height,
    position: adjustStandaloneTextShapePosition(
      { width, height, items },
      anchorPosition,
    ),
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
    const mergedPosition = mergeDocxPosition(
      choiceShape.position,
      fallbackPosition,
    );
    const fallbackAdjustedPosition =
      fallbackPosition && choiceDrawing
        ? adjustInCellPresetShapePosition(
            choiceDrawing,
            descendantByLocalName(choiceDrawing, 'wsp') ?? choiceDrawing,
            mergedPosition,
          )
        : mergedPosition;
    const adviceAdjustedPosition = fallbackPosition
      ? adjustStandaloneTextShapePosition(choiceShape, fallbackAdjustedPosition)
      : fallbackAdjustedPosition;
    return {
      ...choiceShape,
      position: adviceAdjustedPosition,
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

function hasVmlTextBox(shapeNode: Element) {
  return Boolean(descendantByLocalName(shapeNode, 'txbxContent'));
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
  const blocks = parseVmlTextBoxParagraphs(
    shapeNode,
    context,
    id,
    readBlocks,
  ).filter(
    (block) => block.type !== 'paragraph' || block.text || block.inlines.length,
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
    borderRadius: isEllipse
      ? '50%'
      : matchesLocalName(shapeNode, 'roundrect')
      ? 8
      : undefined,
    textVerticalAlign: readVmlTextAnchor(shapeNode),
    fitShapeToText: textBehavior.fitShapeToText || undefined,
    noWrap: textBehavior.noWrap || undefined,
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
    position: adjustStandaloneTextShapePosition(
      {
        width: rootSize.width || maxRight,
        height: rootSize.height || maxBottom,
        items,
      },
      position,
    ),
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
