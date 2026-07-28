import {
  OFFICE_ART_RECORD,
  parseOfficeArtRecords,
  type OfficeArtRecord,
} from '../../../shared/officeart';
import type {
  ChartElement,
  ImageElement,
  ShapeElement,
  SlideBackground,
  SlideElement,
  TextElement,
  ThemeModel,
} from '../../presentation/types';
import { PptRecordReader } from '../binary/PptRecordReader';
import { createPptStaticPreviewCard } from '../images';
import { parsePptTextGroups } from '../text';
import type { PptParseContext } from '../types';
import { readPptOfficeArtProperties } from './readOfficeArtProperties';
import { readPptAnchor } from './readPptAnchor';

const SHAPE_NAMES: Record<number, string> = {
  1: 'rect',
  2: 'roundRect',
  3: 'ellipse',
  4: 'diamond',
  5: 'triangle',
  6: 'rtTriangle',
  7: 'parallelogram',
  8: 'trapezoid',
  20: 'line',
  32: 'line',
  33: 'bentConnector2',
  34: 'curvedConnector2',
};

/** OfficeArt 形状记录中表示对象已删除的标记位。 */
const PPT_SHAPE_FLAG_DELETED = 0x0008;
/** OfficeArt 形状记录中表示对象承担幻灯片背景的标记位。 */
const PPT_SHAPE_FLAG_BACKGROUND = 0x0400;
/** OfficeArt 背景图片填充引用使用的属性编号。 */
const PPT_FILL_BLIP_PROPERTY_ID = 0x0186;
/** OfficeArt 背景纯色填充使用的属性编号。 */
const PPT_FILL_COLOR_PROPERTY_ID = 0x0181;
/** OfficeArt 背景填充透明度使用的 16.16 定点属性编号。 */
const PPT_FILL_OPACITY_PROPERTY_ID = 0x0182;
/** OfficeArt 是否启用填充的布尔属性编号。 */
const PPT_FILL_FLAGS_PROPERTY_ID = 0x01bf;
/** OfficeArt 预设形状的首个几何调节值属性编号。 */
const PPT_GEOMETRY_ADJUST_PROPERTY_ID = 0x0147;
/** OfficeArt 预设形状调节值使用的几何坐标基准。 */
const PPT_GEOMETRY_COORDINATE_SIZE = 21600;
/** OfficeArt 阴影颜色属性编号。 */
const PPT_SHADOW_COLOR_PROPERTY_ID = 0x0201;
/** OfficeArt 阴影透明度使用的 16.16 定点属性编号。 */
const PPT_SHADOW_OPACITY_PROPERTY_ID = 0x0204;
/** OfficeArt 阴影水平偏移属性编号。 */
const PPT_SHADOW_OFFSET_X_PROPERTY_ID = 0x0205;
/** OfficeArt 阴影垂直偏移属性编号。 */
const PPT_SHADOW_OFFSET_Y_PROPERTY_ID = 0x0206;
/** OfficeArt 阴影启用标志属性编号。 */
const PPT_SHADOW_FLAGS_PROPERTY_ID = 0x023f;

/** 描述一页 PPT 绘图解析后的背景和普通元素。 */
export type PptDrawingModel = {
  /** 当前绘图记录声明的幻灯片背景；未声明时交由母版或主题补全。 */
  background?: SlideBackground;
  /** 排除背景记录后可直接交给公共幻灯片渲染器的元素。 */
  elements: SlideElement[];
};

/** 读取 `readColor` 所需的源数据，供PPT 二进制解析使用。 */
function readColor(value: number) {
  const rgb = [value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff];
  return `#${rgb.map((part) => part.toString(16).padStart(2, '0')).join('')}`;
}

/** 查找 `findChild` 对应的目标数据。 */
function findChild(record: OfficeArtRecord, type: number) {
  return record.children?.find((child) => child.type === type);
}

/** 判断 `isBooleanPropertyEnabled` 对应的条件是否成立。 */
function isBooleanPropertyEnabled(value: number, bit: number) {
  const useMask = bit << 16;
  return value & useMask ? Boolean(value & bit) : undefined;
}

/** 将 OfficeArt 16.16 定点透明度转换为标准模型使用的 0 到 1 比例。 */
function readFixedPointOpacity(value: number | undefined) {
  if (value === undefined) return undefined;
  return Math.min(1, Math.max(0, value / 65536));
}

/** 将无符号 OfficeArt 属性还原为有符号 32 位值。 */
function readSignedOfficeArtValue(value: number) {
  return value > 0x7fffffff ? value - 0x100000000 : value;
}

/** 从旧版 PPT 的 OfficeArt 属性恢复卡片等形状的投影效果。 */
function readPptShadow(
  properties: ReturnType<typeof readPptOfficeArtProperties>,
) {
  const flags = properties.get(PPT_SHADOW_FLAGS_PROPERTY_ID)?.value;
  if (flags !== undefined && isBooleanPropertyEnabled(flags, 0x0002) === false)
    return undefined;
  const offsetXValue = properties.get(PPT_SHADOW_OFFSET_X_PROPERTY_ID)?.value;
  const offsetYValue = properties.get(PPT_SHADOW_OFFSET_Y_PROPERTY_ID)?.value;
  const colorValue = properties.get(PPT_SHADOW_COLOR_PROPERTY_ID)?.value;
  const opacityValue = properties.get(PPT_SHADOW_OPACITY_PROPERTY_ID)?.value;
  if (
    offsetXValue === undefined &&
    offsetYValue === undefined &&
    colorValue === undefined &&
    opacityValue === undefined
  )
    return undefined;

  const offsetX = readSignedOfficeArtValue(offsetXValue ?? 0) / 12700;
  const offsetY = readSignedOfficeArtValue(offsetYValue ?? 0) / 12700;
  const distance = Math.hypot(offsetX, offsetY);
  return {
    color: colorValue === undefined ? '#000000' : readColor(colorValue),
    opacity: readFixedPointOpacity(opacityValue),
    offsetX,
    offsetY,
    // 二进制格式未保留柔化半径，按 PowerPoint 对同一偏移投影的默认柔化比例恢复。
    blur: distance * (7 / 3),
  };
}

/** 从无普通锚点的 OfficeArt 背景形状中恢复图片或纯色背景。 */
function parseBackgroundShape(
  record: OfficeArtRecord,
  context: PptParseContext,
): SlideBackground | undefined {
  const fsp = findChild(record, OFFICE_ART_RECORD.FSP);
  if (!fsp || fsp.data.length < 8) return undefined;
  const fspView = new DataView(
    fsp.data.buffer,
    fsp.data.byteOffset,
    fsp.data.byteLength,
  );
  const flags = fspView.getUint32(4, true);
  if (flags & PPT_SHAPE_FLAG_DELETED || !(flags & PPT_SHAPE_FLAG_BACKGROUND)) {
    return undefined;
  }

  const properties = readPptOfficeArtProperties(
    findChild(record, OFFICE_ART_RECORD.FOPT),
  );
  const fillFlags = properties.get(PPT_FILL_FLAGS_PROPERTY_ID)?.value;
  if (
    fillFlags !== undefined &&
    isBooleanPropertyEnabled(fillFlags, 0x0010) === false
  ) {
    return undefined;
  }

  const blipIndex = properties.get(PPT_FILL_BLIP_PROPERTY_ID)?.value;
  const imageRef =
    blipIndex === undefined ? undefined : context.blipUrls.get(blipIndex);
  if (imageRef) return { imageRef };

  const fillColor = properties.get(PPT_FILL_COLOR_PROPERTY_ID)?.value;
  if (fillColor === undefined) return undefined;
  return {
    fill: readColor(fillColor),
    fillOpacity: readFixedPointOpacity(
      properties.get(PPT_FILL_OPACITY_PROPERTY_ID)?.value,
    ),
  };
}

/** 将 OfficeArt 文本锚点转换为公共文本框的垂直对齐方式。 */
function readTextVerticalAlign(value: number | undefined) {
  if (value === 1 || value === 4) return 'middle' as const;
  if ([2, 5, 7, 9].includes(value ?? -1)) return 'bottom' as const;
  return 'top' as const;
}

/** 解析 `parseShape` 接收的数据，并返回PPT 二进制解析结果。 */
function parseShape(
  record: OfficeArtRecord,
  index: number,
  theme: ThemeModel,
  fonts: Map<number, string>,
  context: PptParseContext,
): SlideElement | undefined {
  const fsp = findChild(record, OFFICE_ART_RECORD.FSP);
  const anchor = readPptAnchor(
    findChild(record, OFFICE_ART_RECORD.CLIENT_ANCHOR),
  );
  if (
    !fsp ||
    fsp.data.length < 8 ||
    !anchor ||
    !anchor.width ||
    !anchor.height
  ) {
    return undefined;
  }
  const fspView = new DataView(
    fsp.data.buffer,
    fsp.data.byteOffset,
    fsp.data.byteLength,
  );
  const shapeId = fspView.getUint32(0, true);
  const flags = fspView.getUint32(4, true);
  if (flags & PPT_SHAPE_FLAG_DELETED || flags & PPT_SHAPE_FLAG_BACKGROUND) {
    return undefined;
  }

  const properties = readPptOfficeArtProperties(
    findChild(record, OFFICE_ART_RECORD.FOPT),
  );
  const fillFlags = properties.get(0x01bf)?.value;
  const lineFlags = properties.get(0x01ff)?.value;
  const filled =
    fillFlags === undefined
      ? undefined
      : isBooleanPropertyEnabled(fillFlags, 0x0010);
  const lined =
    lineFlags === undefined
      ? undefined
      : isBooleanPropertyEnabled(lineFlags, 0x0008);
  const shapeType = fsp.instance;
  const shape = SHAPE_NAMES[shapeType] ?? 'rect';
  const fillColor = properties.get(0x0181)?.value;
  const lineColor = properties.get(0x01c0)?.value;
  const lineWidth = properties.get(0x01cb)?.value;
  const rotation = properties.get(0x0004)?.value;
  const adjustValue = properties.get(PPT_GEOMETRY_ADJUST_PROPERTY_ID)?.value;
  const common = {
    id: `ppt-shape-${shapeId}`,
    x: anchor.x,
    y: anchor.y,
    width: anchor.width,
    height: anchor.height,
    rotate: rotation === undefined ? undefined : rotation / 65536,
    flipH: Boolean(flags & 0x0040),
    flipV: Boolean(flags & 0x0080),
    zIndex: index,
    shape,
    shadow: readPptShadow(properties),
    // 二进制 PPT 的圆角调节值以 21600 为基准；0 必须保留，避免误用渲染默认值。
    borderRadius:
      shape === 'roundRect' && adjustValue !== undefined
        ? Math.min(0.5, Math.max(0, adjustValue / PPT_GEOMETRY_COORDINATE_SIZE))
        : undefined,
    fill:
      filled === false
        ? null
        : fillColor === undefined
        ? null
        : readColor(fillColor),
    // 普通形状也必须应用 OfficeArt 透明度，否则半透明遮罩会变成不透明色块。
    fillOpacity:
      filled === false
        ? undefined
        : readFixedPointOpacity(
            properties.get(PPT_FILL_OPACITY_PROPERTY_ID)?.value,
          ),
    stroke:
      lined === false
        ? null
        : lineColor === undefined
        ? undefined
        : readColor(lineColor),
    strokeWidth: lineWidth === undefined ? undefined : lineWidth / 12700,
  };

  const blipIndex = properties.get(0x0104)?.value;
  const clientData = findChild(record, OFFICE_ART_RECORD.CLIENT_DATA);
  let externalObjectId: number | undefined;
  if (clientData) {
    try {
      for (const clientRecord of new PptRecordReader(
        clientData.data,
      ).records()) {
        if (clientRecord.type === 0x0bc1 && clientRecord.data.length >= 4) {
          externalObjectId = new DataView(
            clientRecord.data.buffer,
            clientRecord.data.byteOffset,
            clientRecord.data.byteLength,
          ).getUint32(0, true);
          break;
        }
      }
    } catch {
      // ClientData 损坏时仍可继续使用同一形状的图片预览。
    }
  }
  const embeddedChart = externalObjectId
    ? context.charts.get(externalObjectId)
    : undefined;
  if (embeddedChart) {
    const chart: ChartElement = {
      id: common.id,
      type: 'chart',
      x: common.x,
      y: common.y,
      width: common.width,
      height: common.height,
      rotate: common.rotate,
      zIndex: common.zIndex,
      chart: embeddedChart.chart,
      chartId: `ppt-chart-${externalObjectId}`,
    };
    return chart;
  }
  const imageSource =
    shapeType === 75 && blipIndex ? context.blipUrls.get(blipIndex) : undefined;
  if (imageSource) {
    const image: ImageElement = {
      id: common.id,
      type: 'image',
      x: common.x,
      y: common.y,
      width: common.width,
      height: common.height,
      rotate: common.rotate,
      flipH: common.flipH,
      flipV: common.flipV,
      zIndex: common.zIndex,
      src: imageSource,
      alt: `PowerPoint 图片 ${blipIndex}`,
    };
    return image;
  }
  if (externalObjectId) {
    const image: ImageElement = {
      id: common.id,
      type: 'image',
      x: common.x,
      y: common.y,
      width: common.width,
      height: common.height,
      rotate: common.rotate,
      zIndex: common.zIndex,
      src: createPptStaticPreviewCard(
        '嵌入对象',
        `PowerPoint 对象 ${externalObjectId}`,
        context,
      ),
      alt: `PowerPoint 嵌入对象 ${externalObjectId}`,
    };
    return image;
  }

  const textbox = findChild(record, OFFICE_ART_RECORD.CLIENT_TEXTBOX);
  if (textbox) {
    const textRecords = Array.from(new PptRecordReader(textbox.data).records());
    const groups = parsePptTextGroups(
      textRecords,
      {
        document: {
          fontFamily: theme.fontScheme.minorLatin ?? 'Arial',
          fontSize: 18,
          color: theme.colorScheme.dk1 ?? '#000000',
        },
        fonts,
      },
      context,
    );
    const paragraphs = groups.flatMap((group) => group.paragraphs);
    if (
      paragraphs.some((paragraph) => paragraph.runs.some((run) => run.text))
    ) {
      const textType = groups.find(
        (group) => group.paragraphs.length,
      )?.textType;
      const element: TextElement = {
        ...common,
        type: 'text',
        paragraphs,
        placeholderType:
          textType === 0 || textType === 6
            ? 'title'
            : textType === 1 || textType === 5
            ? 'body'
            : undefined,
        boxStyle: {
          verticalAlign: readTextVerticalAlign(properties.get(0x0087)?.value),
        },
      };
      return element;
    }
  }

  const element: ShapeElement = {
    ...common,
    type: 'shape',
    // null 表示源文件明确禁用填充或线条，只有属性缺省时才使用 Office 默认值。
    fill: common.fill === undefined ? '#ffffff' : common.fill,
    stroke: common.stroke === undefined ? '#000000' : common.stroke,
  };
  return element;
}

/** 提取并汇总 `collectShapeContainers` 返回的数据。 */
function collectShapeContainers(records: OfficeArtRecord[]) {
  const shapes: OfficeArtRecord[] = [];
  const visit = (items: OfficeArtRecord[]) => {
    for (const item of items) {
      if (item.type === OFFICE_ART_RECORD.SP_CONTAINER) shapes.push(item);
      else if (item.children) visit(item.children);
    }
  };
  visit(records);
  return shapes;
}

/** 将一页 PPDrawing 转换为统一的文本与基础图形元素。 */
export function parsePptDrawing(
  bytes: Uint8Array,
  theme: ThemeModel,
  fonts: Map<number, string>,
  context: PptParseContext,
): PptDrawingModel {
  try {
    const records = parseOfficeArtRecords(bytes, context.warnings);
    const shapeContainers = collectShapeContainers(records);
    return {
      background: shapeContainers
        .map((record) => parseBackgroundShape(record, context))
        .find((background) => Boolean(background)),
      elements: shapeContainers
        .map((record, index) =>
          parseShape(record, index, theme, fonts, context),
        )
        .filter((element): element is SlideElement => Boolean(element)),
    };
  } catch (error) {
    context.warnings.push({
      code: 'PPT_DRAWING_CORRUPT',
      message:
        error instanceof Error ? error.message : 'OfficeArt 绘图记录无法读取',
    });
    return { elements: [] };
  }
}
