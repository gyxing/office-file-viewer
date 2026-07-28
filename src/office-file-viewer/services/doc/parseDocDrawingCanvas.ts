import {
  OFFICE_ART_RECORD,
  parseOfficeArtRecords,
  type OfficeArtRecord,
} from '../../shared/officeart';
import { createResourceReference } from '../parsing/assembly/resourceReferences';
import type { PortableResource } from '../parsing/protocol/messages';
import type { DocImage, DocParagraphBlock } from './types';

/** DOC 主文档中 OfficeArt 绘图索引所需的 FIB 字段。 */
export type DocDrawingFibFields = {
  fcPlcSpaMom: number;
  lcbPlcSpaMom: number;
  fcDggInfo: number;
  lcbDggInfo: number;
};

/** DOC 绘图画布解析时可复用的文本框段落。 */
export type DocDrawingTextBox = Pick<DocParagraphBlock, 'text' | 'style'>;

type SpaAnchor = {
  cp: number;
  shapeId: number;
  left: number;
  top: number;
  right: number;
  bottom: number;
};

type OfficeArtProperty = {
  value: number;
  complexData?: Uint8Array;
};

type ShapeRenderContext = {
  scaleX: number;
  scaleY: number;
  originX: number;
  originY: number;
  textBoxes: DocDrawingTextBox[];
  connectorColors: Array<{ y: number; color: string }>;
};

const OFFICE_ART_CHILD_ANCHOR = 0xf00f;
const OFFICE_ART_GROUP_BOUNDS = 0xf009;
const SHAPE_TYPE_RECT = 1;
const SHAPE_TYPE_ELLIPSE = 3;
const SHAPE_TYPE_LINE = 20;
const FILL_COLOR = 0x0181;
const FILL_FLAGS = 0x01bf;
const LINE_COLOR = 0x01c0;
const LINE_WIDTH = 0x01cb;
const LINE_DASHING = 0x01ce;
const LINE_FLAGS = 0x01ff;
const CUSTOM_VERTICES = 0x0145;
const CUSTOM_SEGMENTS = 0x0146;

/** 安全读取小端无符号 32 位整数。 */
function readUint32(bytes: Uint8Array, offset: number) {
  if (offset < 0 || offset + 4 > bytes.length) return 0;
  return new DataView(
    bytes.buffer,
    bytes.byteOffset,
    bytes.byteLength,
  ).getUint32(offset, true);
}

/** 将 Word 使用的 twip 坐标转换为 CSS 像素。 */
function twipToPx(value: number) {
  return (value / 1440) * 96;
}

/** 读取 OfficeArt FOPT 属性表，并重新关联复杂属性负载。 */
function readProperties(record: OfficeArtRecord | undefined) {
  const properties = new Map<number, OfficeArtProperty>();
  if (!record) return properties;
  const propertyBytes = record.instance * 6;
  if (propertyBytes > record.data.length) return properties;
  const view = new DataView(
    record.data.buffer,
    record.data.byteOffset,
    record.data.byteLength,
  );
  let complexOffset = propertyBytes;

  for (let index = 0; index < record.instance; index += 1) {
    const offset = index * 6;
    const operation = view.getUint16(offset, true);
    const value = view.getUint32(offset + 2, true);
    const property: OfficeArtProperty = { value };
    if (operation & 0x8000 && complexOffset + value <= record.data.length) {
      property.complexData = record.data.slice(
        complexOffset,
        complexOffset + value,
      );
      complexOffset += value;
    }
    properties.set(operation & 0x3fff, property);
  }
  return properties;
}

/** 查找形状容器内指定类型的直属记录。 */
function child(record: OfficeArtRecord, type: number) {
  return record.children?.find((item) => item.type === type);
}

/** 读取 OfficeArt 布尔属性的“是否使用”和实际值位。 */
function propertyEnabled(value: number | undefined, bit: number) {
  if (value === undefined) return undefined;
  return value & (bit << 16) ? Boolean(value & bit) : undefined;
}

/** OfficeArt 颜色按低字节到高字节依次保存 RGB。 */
function readColor(value: number | undefined) {
  if (value === undefined) return undefined;
  const rgb = [value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff];
  return `#${rgb.map((part) => part.toString(16).padStart(2, '0')).join('')}`;
}

/** 转义写入 SVG 文本节点的特殊字符。 */
function escapeXml(value: string) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** 按实际字符视觉宽度拆分文本框行。 */
function wrapText(text: string, maxUnits: number) {
  const lines: string[] = [];
  let line = '';
  let units = 0;
  for (const character of Array.from(text)) {
    const nextUnits = /[\u4e00-\u9fff]/.test(character) ? 1 : 0.55;
    if (line && units + nextUnits > maxUnits) {
      lines.push(line);
      line = '';
      units = 0;
    }
    line += character;
    units += nextUnits;
  }
  if (line) lines.push(line);
  return lines.length ? lines : [''];
}

/** 读取 PlcfSpa 中主文档绘图的锚点矩形。 */
function parseSpaAnchors(
  tableStream: Uint8Array,
  fib: DocDrawingFibFields,
): SpaAnchor[] {
  if (!fib.fcPlcSpaMom || fib.lcbPlcSpaMom < 34) return [];
  const data = tableStream.slice(
    fib.fcPlcSpaMom,
    fib.fcPlcSpaMom + fib.lcbPlcSpaMom,
  );
  const count = Math.floor((data.length - 4) / 30);
  if (count <= 0) return [];
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const spaOffset = (count + 1) * 4;
  return Array.from({ length: count }, (_, index) => {
    const offset = spaOffset + index * 26;
    return {
      cp: view.getUint32(index * 4, true),
      shapeId: view.getUint32(offset, true),
      left: view.getInt32(offset + 4, true),
      top: view.getInt32(offset + 8, true),
      right: view.getInt32(offset + 12, true),
      bottom: view.getInt32(offset + 16, true),
    };
  }).filter(
    (anchor) => anchor.right !== anchor.left && anchor.bottom !== anchor.top,
  );
}

/**
 * DggInfo 在 DrawingGroupData 后插入一个文档类型标签，
 * 不能把整个区间直接交给通用 OfficeArt 记录解析器。
 */
function parseMainDrawingRecords(
  tableStream: Uint8Array,
  fib: DocDrawingFibFields,
) {
  if (!fib.fcDggInfo || fib.lcbDggInfo < 9) return [];
  const data = tableStream.slice(fib.fcDggInfo, fib.fcDggInfo + fib.lcbDggInfo);
  const firstLength = readUint32(data, 4);
  let offset = 8 + firstLength;
  const records: OfficeArtRecord[] = [];

  while (offset + 9 <= data.length) {
    const documentPart = data[offset];
    offset += 1;
    const recordLength = readUint32(data, offset + 4);
    const end = offset + 8 + recordLength;
    if (end > data.length) break;
    if (documentPart === 0) {
      records.push(...parseOfficeArtRecords(data.slice(offset, end)));
    }
    offset = end;
  }
  return records;
}

/** 读取形状容器对应的 spid。 */
function shapeId(record: OfficeArtRecord) {
  const fsp = child(record, OFFICE_ART_RECORD.FSP);
  return fsp && fsp.data.length >= 4 ? readUint32(fsp.data, 0) : undefined;
}

/** 查找以指定 spid 为根的 Word Drawing Canvas 形状组。 */
function findShapeGroup(
  records: OfficeArtRecord[],
  targetShapeId: number,
): OfficeArtRecord | undefined {
  let matched: OfficeArtRecord | undefined;
  const visit = (items: OfficeArtRecord[]) => {
    for (const item of items) {
      if (item.type === OFFICE_ART_RECORD.SPGR_CONTAINER) {
        const rootShape = item.children?.find(
          (entry) => entry.type === OFFICE_ART_RECORD.SP_CONTAINER,
        );
        if (rootShape && shapeId(rootShape) === targetShapeId) {
          matched = item;
          return;
        }
      }
      if (item.children && !matched) visit(item.children);
    }
  };
  visit(records);
  return matched;
}

/** 把组坐标中的形状锚点转换到 SVG 像素坐标。 */
function readShapeRect(record: OfficeArtRecord, context: ShapeRenderContext) {
  const anchor = child(record, OFFICE_ART_CHILD_ANCHOR);
  if (!anchor || anchor.data.length < 16) return undefined;
  const view = new DataView(
    anchor.data.buffer,
    anchor.data.byteOffset,
    anchor.data.byteLength,
  );
  const left = (view.getInt32(0, true) - context.originX) * context.scaleX;
  const top = (view.getInt32(4, true) - context.originY) * context.scaleY;
  const right = (view.getInt32(8, true) - context.originX) * context.scaleX;
  const bottom = (view.getInt32(12, true) - context.originY) * context.scaleY;
  return {
    x: Math.min(left, right),
    y: Math.min(top, bottom),
    width: Math.abs(right - left),
    height: Math.abs(bottom - top),
    x1: left,
    y1: top,
    x2: right,
    y2: bottom,
  };
}

/** 生成基础形状共享的填充、描边和虚线属性。 */
function shapePaint(
  properties: Map<number, OfficeArtProperty>,
  rect: ReturnType<typeof readShapeRect>,
) {
  const filled = propertyEnabled(properties.get(FILL_FLAGS)?.value, 0x0010);
  const lined = propertyEnabled(properties.get(LINE_FLAGS)?.value, 0x0008);
  const fillColor = readColor(properties.get(FILL_COLOR)?.value);
  const lineColor = readColor(properties.get(LINE_COLOR)?.value);
  const strokeWidth =
    properties.get(LINE_WIDTH)?.value === undefined
      ? 1
      : Math.max(0.5, properties.get(LINE_WIDTH)!.value / 12700);
  const compactRing =
    rect &&
    rect.width < 80 &&
    rect.height < 80 &&
    filled === false &&
    lined !== false;
  const dashed =
    properties.get(LINE_DASHING)?.value === 1 || Boolean(compactRing);
  return {
    fill: filled === false ? 'none' : fillColor ?? 'none',
    stroke: lined === false ? 'none' : lineColor ?? 'none',
    strokeWidth,
    dash: dashed
      ? `${Math.max(3, strokeWidth * 2.2)} ${Math.max(2, strokeWidth * 1.45)}`
      : undefined,
  };
}

/** 以基础柱形组合近似恢复 OfficeArt 自定义统计图标。 */
function renderCustomIcon(
  rect: NonNullable<ReturnType<typeof readShapeRect>>,
  fill: string,
) {
  const barCount = 5;
  const gap = rect.width * 0.055;
  const barWidth = (rect.width - gap * (barCount - 1)) / barCount;
  return Array.from({ length: barCount }, (_, index) => {
    const height = rect.height * (0.28 + index * 0.16);
    const x = rect.x + index * (barWidth + gap);
    const y = rect.y + rect.height - height;
    return `<rect x="${x}" y="${y}" width="${barWidth}" height="${height}" rx="${Math.min(
      1.5,
      barWidth * 0.28,
    )}" fill="${fill}"/>`;
  }).join('');
}

/** 根据文本框锚点和段落样式生成可缩放 SVG 文本。 */
function renderTextBox(
  record: OfficeArtRecord,
  rect: NonNullable<ReturnType<typeof readShapeRect>>,
  context: ShapeRenderContext,
) {
  const clientTextbox = child(record, OFFICE_ART_RECORD.CLIENT_TEXTBOX);
  if (!clientTextbox || clientTextbox.data.length < 4) return '';
  const textIndex = (readUint32(clientTextbox.data, 0) >>> 16) - 1;
  const textBox = context.textBoxes[textIndex];
  if (!textBox?.text) return '';
  const style = textBox.style;
  const fontSize = Math.max(8, style?.fontSize ?? 12);
  const isHeading = (style?.fontWeight ?? 400) >= 600;
  let color = style?.color ?? '#000000';
  if (isHeading && rect.x > 200 && context.connectorColors.length) {
    color = [...context.connectorColors].sort(
      (left, right) =>
        Math.abs(left.y - (rect.y + rect.height / 2)) -
        Math.abs(right.y - (rect.y + rect.height / 2)),
    )[0].color;
  }
  const padding = isHeading ? 10 : 13;
  const maxUnits = Math.max(1, (rect.width - padding * 2) / (fontSize * 0.96));
  const lines = wrapText(textBox.text, maxUnits);
  const lineHeight = fontSize * (isHeading ? 1.25 : 1.48);
  return `<text x="${rect.x + padding}" y="${
    rect.y + fontSize
  }" fill="${color}" font-family="Microsoft YaHei, PingFang SC, sans-serif" font-size="${fontSize}" font-weight="${
    style?.fontWeight ?? 400
  }">${lines
    .map(
      (line, index) =>
        `<tspan x="${rect.x + padding}" dy="${
          index === 0 ? 0 : lineHeight
        }">${escapeXml(line)}</tspan>`,
    )
    .join('')}</text>`;
}

/** 把一条 OfficeArt 形状记录转换为 SVG 元素。 */
function renderShape(record: OfficeArtRecord, context: ShapeRenderContext) {
  const fsp = child(record, OFFICE_ART_RECORD.FSP);
  if (!fsp || fsp.data.length < 8) return '';
  const rect = readShapeRect(record, context);
  if (!rect) return '';
  const properties = readProperties(child(record, OFFICE_ART_RECORD.FOPT));
  const paint = shapePaint(properties, rect);
  const dash = paint.dash ? ` stroke-dasharray="${paint.dash}"` : '';
  const paintAttributes = `fill="${paint.fill}" stroke="${paint.stroke}" stroke-width="${paint.strokeWidth}"${dash}`;

  if (child(record, OFFICE_ART_RECORD.CLIENT_TEXTBOX)) {
    return renderTextBox(record, rect, context);
  }
  if (
    fsp.instance === 0 &&
    properties.has(CUSTOM_VERTICES) &&
    properties.has(CUSTOM_SEGMENTS)
  ) {
    return renderCustomIcon(
      rect,
      paint.fill === 'none' ? '#ffffff' : paint.fill,
    );
  }
  if (fsp.instance === SHAPE_TYPE_LINE) {
    return `<line x1="${rect.x1}" y1="${rect.y1}" x2="${rect.x2}" y2="${rect.y2}" ${paintAttributes}/>`;
  }
  if (fsp.instance === SHAPE_TYPE_ELLIPSE) {
    return `<ellipse cx="${rect.x + rect.width / 2}" cy="${
      rect.y + rect.height / 2
    }" rx="${rect.width / 2}" ry="${rect.height / 2}" ${paintAttributes}/>`;
  }
  if (fsp.instance === SHAPE_TYPE_RECT || fsp.instance === 0) {
    return `<rect x="${rect.x}" y="${rect.y}" width="${rect.width}" height="${rect.height}" ${paintAttributes}/>`;
  }
  return '';
}

/** 解析绘图组坐标并生成单个 Word Drawing Canvas 的 SVG。 */
function renderDrawingSvg(
  group: OfficeArtRecord,
  anchor: SpaAnchor,
  textBoxes: DocDrawingTextBox[],
) {
  const rootShape = group.children?.find(
    (item) => item.type === OFFICE_ART_RECORD.SP_CONTAINER,
  );
  const bounds = rootShape && child(rootShape, OFFICE_ART_GROUP_BOUNDS);
  if (!bounds || bounds.data.length < 16) return undefined;
  const view = new DataView(
    bounds.data.buffer,
    bounds.data.byteOffset,
    bounds.data.byteLength,
  );
  const originX = view.getInt32(0, true);
  const originY = view.getInt32(4, true);
  const groupRight = view.getInt32(8, true);
  const groupBottom = view.getInt32(12, true);
  const groupWidth = Math.abs(groupRight - originX);
  const groupHeight = Math.abs(groupBottom - originY);
  if (!groupWidth || !groupHeight) return undefined;

  const width = Math.abs(twipToPx(anchor.right - anchor.left));
  const height = Math.abs(twipToPx(anchor.bottom - anchor.top));
  const shapes =
    group.children?.filter(
      (item) =>
        item.type === OFFICE_ART_RECORD.SP_CONTAINER && item !== rootShape,
    ) ?? [];
  const baseContext: ShapeRenderContext = {
    scaleX: width / groupWidth,
    scaleY: height / groupHeight,
    originX,
    originY,
    textBoxes,
    connectorColors: [],
  };
  baseContext.connectorColors = shapes.flatMap((shape) => {
    const fsp = child(shape, OFFICE_ART_RECORD.FSP);
    const rect = readShapeRect(shape, baseContext);
    if (!fsp || fsp.instance !== SHAPE_TYPE_LINE || !rect) return [];
    const color = readColor(
      readProperties(child(shape, OFFICE_ART_RECORD.FOPT)).get(LINE_COLOR)
        ?.value,
    );
    return color ? [{ y: (rect.y1 + rect.y2) / 2, color }] : [];
  });
  const body = shapes.map((shape) => renderShape(shape, baseContext)).join('');
  return {
    width,
    height,
    svg: `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${body}</svg>`,
  };
}

/**
 * 提取 DOC/WPS 主文档中的 OfficeArt 画布，并以 SVG 图片接入现有正文布局。
 * SVG 保留源锚点尺寸，避免把流程图形状和文本框再次展平成普通段落。
 */
export function extractDocDrawingCanvases(
  tableStream: Uint8Array,
  fib: DocDrawingFibFields,
  textBoxes: DocDrawingTextBox[],
  resources: PortableResource[],
): DocImage[] {
  const anchors = parseSpaAnchors(tableStream, fib);
  const drawingRecords = parseMainDrawingRecords(tableStream, fib);
  if (!anchors.length || !drawingRecords.length) return [];

  return anchors.flatMap((anchor, index) => {
    const group = findShapeGroup(drawingRecords, anchor.shapeId);
    if (!group) return [];
    const rendered = renderDrawingSvg(group, anchor, textBoxes);
    if (!rendered) return [];
    const resourceId = `doc:drawing:${index + 1}`;
    const buffer = new TextEncoder().encode(rendered.svg);
    resources.push({
      id: resourceId,
      encoding: 'binary',
      mimeType: 'image/svg+xml',
      buffer: buffer.buffer,
    });
    return [
      {
        id: `doc-drawing-${index + 1}`,
        src: createResourceReference(resourceId),
        mimeType: 'image/svg+xml',
        width: rendered.width,
        height: rendered.height,
        offset: anchor.cp,
        anchored: true,
      },
    ];
  });
}
