import type { EChartsOption } from 'echarts';
import {
  DEFAULT_OFFICE_THEME,
  resolveOfficeThemeColor,
  type OfficeTheme,
} from './theme';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  descendantByLocalName,
  descendantsByLocalName,
  parseXml,
  textContent,
} from './xml';

/** 描述 OfficeChartType 在 OOXML 公共解析中的数据结构。 */
export type OfficeChartType =
  | 'line'
  | 'bar'
  | 'column'
  | 'pie'
  | 'doughnut'
  | 'area'
  | 'scatter'
  | 'bubble'
  | 'radar'
  | 'map'
  | 'unknown';

/** 描述 OfficeChartSeries 在 OOXML 公共解析中的数据结构。 */
export type OfficeChartSeries = {
  /** OfficeChartSeries 的可读名称。 */
  name: string;
  /** OfficeChartSeries 包含的 values 有序集合。 */
  values: number[];
  /** OfficeChartSeries 包含的 xValues 有序集合。 */
  xValues?: number[];
  /** OfficeChartSeries 包含的 bubbleSizes 有序集合。 */
  bubbleSizes?: number[];
  /** 用于区分 OfficeChartSeries 不同结构分支的类型标识。 */
  type?: OfficeChartType;
  /** OfficeChartSeries 关联的 stacking 结构；字段形状由 'stacked' | 'percentStacked' 定义；未提供时使用来源格式或渲染器的默认行为。 */
  stacking?: 'stacked' | 'percentStacked';
  /** OfficeChartSeries 的 stackGroup 文本值。 */
  stackGroup?: string;
  /** OfficeChartSeries 的 gapWidth 图表布局参数，数值语义遵循 Office 图表规范；未提供时使用来源格式或渲染器的默认行为。 */
  gapWidth?: number;
  /** OfficeChartSeries 的 overlap 图表布局参数，数值语义遵循 Office 图表规范；未提供时使用来源格式或渲染器的默认行为。 */
  overlap?: number;
  /** OfficeChartSeries 的前景或文本颜色，使用标准化 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
  color?: string;
  /** OfficeChartSeries 包含的 pointColors 有序集合。 */
  pointColors?: string[];
  /** OfficeChartSeries 包含的 pointLabels 有序集合。 */
  pointLabels?: string[];
  /** OfficeChartSeries 包含的 pointStyles 有序集合。 */
  pointStyles?: Array<{
    /** OfficeChartSeries 的前景或文本颜色，使用标准化 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
    color?: OfficeChartColor;
    /** OfficeChartSeries 的 borderColor 文本值。 */
    borderColor?: string;
    /** OfficeChartSeries 的 borderWidth 渲染数值，单位为标准化渲染像素。 */
    borderWidth?: number;
  }>;
  /** OfficeChartSeries 的图表数据标签显示配置；未提供时使用来源格式或渲染器的默认行为。 */
  dataLabels?: OfficeDataLabels;
  /** 是否使用平滑曲线连接数据点；未提供时使用来源格式或渲染器的默认行为。 */
  smooth?: boolean;
  /** OfficeChartSeries 的 lineWidth 渲染尺寸，单位为标准化像素；未提供时使用来源格式或渲染器的默认行为。 */
  lineWidth?: number;
  /** 数据系列的数据点标记样式；未提供时不绘制标记。 */
  marker?: {
    /** OfficeChartSeries 的 symbol 文本值。 */
    symbol?: string;
    /** OfficeChartSeries 的 size 数值；具体语义遵循对应源文件格式；未提供时使用来源格式或渲染器的默认行为。 */
    size?: number;
  };
};

/** 描述 OfficeDataLabels 在 OOXML 公共解析中的数据结构。 */
export type OfficeDataLabels = {
  /** 是否按源文件指示隐藏当前图表元素；未提供时使用来源格式或渲染器的默认行为。 */
  delete?: boolean;
  /** OfficeDataLabels 的定位信息及其参考坐标系。 */
  position?: string;
  /** OfficeDataLabels 的 separator 文本值。 */
  separator?: string;
  /** 是否显示 LegendKey 对应的图表或界面元素；未提供时使用来源格式或渲染器的默认行为。 */
  showLegendKey?: boolean;
  /** 是否显示 Val 对应的图表或界面元素；未提供时使用来源格式或渲染器的默认行为。 */
  showVal?: boolean;
  /** 是否显示 CatName 对应的图表或界面元素；未提供时使用来源格式或渲染器的默认行为。 */
  showCatName?: boolean;
  /** 是否显示 SerName 对应的图表或界面元素；未提供时使用来源格式或渲染器的默认行为。 */
  showSerName?: boolean;
  /** 是否显示 Percent 对应的图表或界面元素；未提供时使用来源格式或渲染器的默认行为。 */
  showPercent?: boolean;
  /** 是否显示 BubbleSize 对应的图表或界面元素；未提供时使用来源格式或渲染器的默认行为。 */
  showBubbleSize?: boolean;
  /** 是否显示 LeaderLines 对应的图表或界面元素；未提供时使用来源格式或渲染器的默认行为。 */
  showLeaderLines?: boolean;
};

/** 描述 OOXML 公共解析使用的标准化模型。 */
export type OfficeChartModel = {
  /** 用于区分 OfficeChartModel 不同结构分支的类型标识。 */
  type: OfficeChartType;
  /** OfficeChartModel 对外展示的标题。 */
  title?: string;
  /** OfficeChartModel 包含的 categories 有序集合。 */
  categories: string[];
  /** OfficeChartModel 包含的 series 有序集合。 */
  series: OfficeChartSeries[];
  /** OfficeChartModel 的图表数据标签显示配置；未提供时使用来源格式或渲染器的默认行为。 */
  dataLabels?: OfficeDataLabels;
  /** 是否显示图表图例；未提供时使用来源格式或渲染器的默认行为。 */
  showLegend?: boolean;
  /** OfficeChartModel 的图例停靠位置；未提供时使用来源格式或渲染器的默认行为。 */
  legendPosition?: 'top' | 'bottom' | 'left' | 'right';
  /** OfficeChartModel 的图例尺寸和文字样式；未提供时使用来源格式或渲染器的默认行为。 */
  legendStyle?: {
    /** OfficeChartModel 的 itemWidth 渲染尺寸，单位为标准化像素；未提供时使用来源格式或渲染器的默认行为。 */
    itemWidth?: number;
    /** OfficeChartModel 的 itemHeight 渲染尺寸，单位为标准化像素；未提供时使用来源格式或渲染器的默认行为。 */
    itemHeight?: number;
    /** 图例文本的字体与颜色样式；未提供时沿用图表主题。 */
    textStyle?: {
      /** OfficeChartModel 的前景或文本颜色，使用标准化 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
      color?: string;
      /** OfficeChartModel 的字体族名称；未提供时沿用来源格式或渲染器的默认规则。 */
      fontFamily?: string;
      /** OfficeChartModel 的字号，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
      fontSize?: number;
      /** OfficeChartModel 的字体样式；未提供时沿用来源格式或渲染器的默认规则。 */
      fontStyle?: 'normal' | 'italic' | 'oblique';
      /** OfficeChartModel 的字体粗细值；未提供时沿用来源格式或渲染器的默认规则。 */
      fontWeight?: 'normal' | 'bold' | 'bolder' | 'lighter' | number;
    };
  };
  /** 是否显示图表数据标签；未提供时使用来源格式或渲染器的默认行为。 */
  showDataLabels?: boolean;
  /** OfficeChartModel 的 holeSize 图表布局参数，数值语义遵循 Office 图表规范；未提供时使用来源格式或渲染器的默认行为。 */
  holeSize?: number;
  /** OfficeChartModel 的旋转或起始角度，单位为度；未提供时使用来源格式或渲染器的默认行为。 */
  startAngle?: number;
  /** OfficeChartModel 关联的 ofPieType 结构；字段形状由 'bar' | 'pie' 定义；未提供时使用来源格式或渲染器的默认行为。 */
  ofPieType?: 'bar' | 'pie';
  /** OfficeChartModel 对应项目的数量。 */
  ofPieSecondPlotCount?: number;
  /** OfficeChartModel 的 secondPieSize 图表布局参数，数值语义遵循 Office 图表规范；未提供时使用来源格式或渲染器的默认行为。 */
  secondPieSize?: number;
  /** OfficeChartModel 的 gapWidth 图表布局参数，数值语义遵循 Office 图表规范；未提供时使用来源格式或渲染器的默认行为。 */
  gapWidth?: number;
  /** OfficeChartModel 的 overlap 图表布局参数，数值语义遵循 Office 图表规范；未提供时使用来源格式或渲染器的默认行为。 */
  overlap?: number;
  /** OfficeChartModel 关联的 roseType 结构；字段形状由 'radius' | 'area' 定义；未提供时使用来源格式或渲染器的默认行为。 */
  roseType?: 'radius' | 'area';
  /** OfficeChartModel 关联的 radius 结构；字段形状由 [string, string] 定义；未提供时使用来源格式或渲染器的默认行为。 */
  radius?: [string, string];
  /** OfficeChartModel 的 radarStyle 文本值。 */
  radarStyle?: string;
  /** OfficeChartModel 的 radarRadius 文本值。 */
  radarRadius?: string;
  /** OfficeChartModel 的旋转或起始角度，单位为度；未提供时使用来源格式或渲染器的默认行为。 */
  radarStartAngle?: number;
  /** OfficeChartModel 的 radarSplitNumber 数值；具体语义遵循对应源文件格式；未提供时使用来源格式或渲染器的默认行为。 */
  radarSplitNumber?: number;
  /** OfficeChartModel 包含的 radarIndicators 有序集合。 */
  radarIndicators?: Array<{
    /** OfficeChartModel 的可读名称。 */
    name: string;
    /** OfficeChartModel 的 max 数值；具体语义遵循对应源文件格式。 */
    max: number;
  }>;
  /** OfficeChartModel 的 mapSeriesName 文本值。 */
  mapSeriesName?: string;
  /** OfficeChartModel 的 mapRegion 文本值。 */
  mapRegion?: string;
  /** OfficeChartModel 的 mapName 文本值。 */
  mapName?: string;
  /** OfficeChartModel 的 mapGeoJsonUrl 文本值。 */
  mapGeoJsonUrl?: string;
  /** OfficeChartModel 的 snapshotSrc 文本值。 */
  snapshotSrc?: string;
  /** OfficeChartModel 的 sourceType 文本值。 */
  sourceType?: string;
  /** OfficeChartModel 关联的 renderMode 结构；字段形状由 'interactive' | 'snapshot' 定义；未提供时使用来源格式或渲染器的默认行为。 */
  renderMode?: 'interactive' | 'snapshot';
  /** OfficeChartModel 的 degradedFrom 文本值。 */
  degradedFrom?: string;
};

const DEFAULT_COLORS = [
  '#5470c6',
  '#91cc75',
  '#fac858',
  '#ee6666',
  '#73c0de',
  '#3ba272',
  '#fc8452',
];
const OFFICE_FONT_FAMILY =
  '"Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", Arial, sans-serif';
const OFFICE_TEXT_STYLE = {
  color: '#334155',
  fontFamily: OFFICE_FONT_FAMILY,
};

/** 描述 OfficeChartColorStop 在 OOXML 公共解析中的数据结构。 */
type OfficeChartColorStop = {
  /** OfficeChartColorStop 在所属数据范围中的偏移位置。 */
  offset: number;
  /** OfficeChartColorStop 的前景或文本颜色，使用标准化 CSS 颜色值。 */
  color: string;
};

/** 描述 OfficeChartColor 在 OOXML 公共解析中的数据结构。 */
type OfficeChartColor =
  | string
  | {
      /** 用于区分 OfficeChartColor 不同结构分支的类型标识。 */
      type: 'linear';
      /** OfficeChartColor 的 x 尺寸或坐标，单位为标准化渲染像素。 */
      x: number;
      /** OfficeChartColor 的 y 尺寸或坐标，单位为标准化渲染像素。 */
      y: number;
      /** OfficeChartColor 在矢量图元坐标系中的 x2 几何值。 */
      x2: number;
      /** OfficeChartColor 在矢量图元坐标系中的 y2 几何值。 */
      y2: number;
      /** OfficeChartColor 包含的 colorStops 有序集合。 */
      colorStops: OfficeChartColorStop[];
      /** 渐变坐标是否相对于全局画布计算；未提供时使用来源格式或渲染器的默认行为。 */
      global?: boolean;
    };

const CHART_NODE_TO_TYPE: Record<string, OfficeChartType> = {
  linechart: 'line',
  barchart: 'column',
  piechart: 'pie',
  doughnutchart: 'doughnut',
  areachart: 'area',
  scatterchart: 'scatter',
  bubblechart: 'bubble',
  radarchart: 'radar',
  ofpiechart: 'pie',
};

/** 解码 `decodeMojibake` 接收的源数据。 */
export function decodeMojibake(value: string) {
  if (!/[脙脗盲氓忙莽猫茅]|锟|鍥|绯|绫|诲|埆|垪|棰/.test(value)) {
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

/** 执行 `firstText` 封装的 OOXML 公共解析处理步骤。 */
function firstText(node: Element | null | undefined) {
  const value =
    textContent(descendantByLocalName(node, 't')) ||
    textContent(descendantByLocalName(node, 'v'));
  return decodeMojibake(value.trim());
}

/** 读取 `readCacheValues` 所需的源数据，供 OOXML 公共解析使用。 */
function readCacheValues(node: Element | null | undefined, date1904 = false) {
  const strCache = descendantByLocalName(node, 'strCache');
  if (strCache) {
    return descendantsByLocalName(strCache, 'pt')
      .sort((a, b) => Number(attr(a, 'idx') ?? 0) - Number(attr(b, 'idx') ?? 0))
      .map((point) =>
        decodeMojibake(textContent(childByLocalName(point, 'v')).trim()),
      );
  }

  const numCache = descendantByLocalName(node, 'numCache');
  if (!numCache) return [];

  const cacheFormatCode = decodeMojibake(
    textContent(childByLocalName(numCache, 'formatCode')).trim(),
  );
  return descendantsByLocalName(numCache, 'pt')
    .sort((a, b) => Number(attr(a, 'idx') ?? 0) - Number(attr(b, 'idx') ?? 0))
    .map((point) => {
      const value = decodeMojibake(
        textContent(childByLocalName(point, 'v')).trim(),
      );
      const formatCode = decodeMojibake(
        attr(point, 'formatCode') ?? cacheFormatCode,
      );
      return formatCacheValue(value, formatCode, date1904);
    });
}

/** 读取 `readNumericValues` 所需的源数据，供 OOXML 公共解析使用。 */
function readNumericValues(node: Element | null | undefined) {
  return readCacheValues(node)
    .map((value) => Number(value))
    .filter((value) => Number.isFinite(value));
}

/** 将输入标准化为 `normalizeType` 返回的结构。 */
function normalizeType(chartNode: Element | null): OfficeChartType {
  if (!chartNode) return 'unknown';
  const localName = (
    chartNode.localName.split(':').pop() ?? chartNode.localName
  ).toLowerCase();
  if (localName === 'barchart') {
    const barDir = attr(childByLocalName(chartNode, 'barDir'), 'val');
    return barDir === 'bar' ? 'bar' : 'column';
  }
  return CHART_NODE_TO_TYPE[localName] ?? 'unknown';
}

/** 读取 `readSeriesColorWithTheme` 所需的源数据，供 OOXML 公共解析使用。 */
function readSeriesColorWithTheme(seriesNode: Element, theme: OfficeTheme) {
  const spPr = childByLocalName(seriesNode, 'spPr');
  const fillNode =
    childByLocalName(spPr, 'solidFill') ?? childByLocalName(spPr, 'gradFill');
  const lineNode = childByLocalName(spPr, 'ln');
  const color =
    readFillValue(fillNode, theme) ?? readFillColor(lineNode, theme);
  if (typeof color === 'string') return color;

  const fallbackFill = childByLocalName(
    childByLocalName(seriesNode, 'spPr'),
    'solidFill',
  );
  const fallbackLine = childByLocalName(
    childByLocalName(seriesNode, 'spPr'),
    'ln',
  );
  const fallback =
    readFillValue(fallbackFill, DEFAULT_OFFICE_THEME) ??
    readFillColor(fallbackLine, DEFAULT_OFFICE_THEME);
  return typeof fallback === 'string' ? fallback : undefined;
}

/** 读取 `readFillColor` 所需的源数据，供 OOXML 公共解析使用。 */
function readFillColor(node: Element | null | undefined, theme: OfficeTheme) {
  const fillNode =
    childByLocalName(node, 'solidFill') ?? childByLocalName(node, 'gradFill');
  const color = readFillValue(fillNode, theme);
  return typeof color === 'string' ? color : undefined;
}

/** 读取 `readPointStyles` 所需的源数据，供 OOXML 公共解析使用。 */
function readPointStyles(seriesNode: Element, theme: OfficeTheme) {
  const styles: OfficeChartSeries['pointStyles'] = [];
  childrenByLocalName(seriesNode, 'dPt').forEach((pointNode) => {
    const index = Number(attr(childByLocalName(pointNode, 'idx'), 'val'));
    if (!Number.isFinite(index)) return;
    const spPr = childByLocalName(pointNode, 'spPr');
    if (!spPr) return;
    const fillNode =
      childByLocalName(spPr, 'solidFill') ?? childByLocalName(spPr, 'gradFill');
    const lineNode = childByLocalName(spPr, 'ln');
    const color = readFillValue(fillNode, theme);
    const borderColor = readFillColor(lineNode, theme);
    const borderWidth = readLineWidth(lineNode);
    const hasVisibleBorder = Boolean(
      lineNode &&
        !childByLocalName(lineNode, 'noFill') &&
        (borderColor || borderWidth !== undefined),
    );
    if (color || hasVisibleBorder) {
      styles[index] = {
        color,
        borderColor,
        borderWidth: borderWidth ?? (hasVisibleBorder ? 1 : undefined),
      };
    }
  });
  return styles;
}

/** 读取 `readPointColors` 所需的源数据，供 OOXML 公共解析使用。 */
function readPointColors(seriesNode: Element, theme: OfficeTheme) {
  const colors: string[] = [];
  childrenByLocalName(seriesNode, 'dPt').forEach((pointNode) => {
    const index = Number(attr(childByLocalName(pointNode, 'idx'), 'val'));
    if (!Number.isFinite(index)) return;
    const color = readFillColor(childByLocalName(pointNode, 'spPr'), theme);
    if (color) colors[index] = color;
  });
  return colors;
}

/** 执行 `localName` 封装的 OOXML 公共解析处理步骤。 */
function localName(node: Element | null | undefined) {
  return (
    node?.localName.split(':').pop() ??
    node?.localName ??
    ''
  ).toLowerCase();
}

/** 执行 `clamp01` 封装的 OOXML 公共解析处理步骤。 */
function clamp01(value: number) {
  return Math.max(0, Math.min(1, value));
}

/** 执行 `clamp255` 封装的 OOXML 公共解析处理步骤。 */
function clamp255(value: number) {
  return Math.max(0, Math.min(255, value));
}

/** 将输入标准化为 `normalizeHex` 返回的结构。 */
function normalizeHex(value?: string) {
  if (!value) return undefined;
  if (/^#?[0-9a-f]{6}$/i.test(value)) {
    return value.startsWith('#') ? value : `#${value}`;
  }
  return undefined;
}

/** 执行 `hexToRgb` 封装的 OOXML 公共解析处理步骤。 */
function hexToRgb(hex: string) {
  const normalized = hex.replace('#', '');
  const value = Number.parseInt(normalized, 16);
  return {
    r: (value >> 16) & 255,
    g: (value >> 8) & 255,
    b: value & 255,
  };
}

/** 执行 `rgbToHex` 封装的 OOXML 公共解析处理步骤。 */
function rgbToHex(r: number, g: number, b: number) {
  return `#${[r, g, b]
    .map((value) => clamp255(value).toString(16).padStart(2, '0'))
    .join('')}`;
}

/** 执行 `rgbToHsl` 封装的 OOXML 公共解析处理步骤。 */
function rgbToHsl(r: number, g: number, b: number) {
  const red = r / 255;
  const green = g / 255;
  const blue = b / 255;
  const max = Math.max(red, green, blue);
  const min = Math.min(red, green, blue);
  const lightness = (max + min) / 2;
  if (max === min) {
    return { h: 0, s: 0, l: lightness };
  }
  const delta = max - min;
  const saturation =
    lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue = 0;
  switch (max) {
    case red:
      hue = (green - blue) / delta + (green < blue ? 6 : 0);
      break;
    case green:
      hue = (blue - red) / delta + 2;
      break;
    default:
      hue = (red - green) / delta + 4;
      break;
  }
  return { h: hue * 60, s: saturation, l: lightness };
}

/** 执行 `hslToRgb` 封装的 OOXML 公共解析处理步骤。 */
function hslToRgb(h: number, s: number, l: number) {
  const hue = (((h % 360) + 360) % 360) / 360;
  if (s === 0) {
    const value = Math.round(l * 255);
    return { r: value, g: value, b: value };
  }

  const hue2rgb = (p: number, q: number, t: number) => {
    let temp = t;
    if (temp < 0) temp += 1;
    if (temp > 1) temp -= 1;
    if (temp < 1 / 6) return p + (q - p) * 6 * temp;
    if (temp < 1 / 2) return q;
    if (temp < 2 / 3) return p + (q - p) * (2 / 3 - temp) * 6;
    return p;
  };

  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return {
    r: Math.round(hue2rgb(p, q, hue + 1 / 3) * 255),
    g: Math.round(hue2rgb(p, q, hue) * 255),
    b: Math.round(hue2rgb(p, q, hue - 1 / 3) * 255),
  };
}

/** 读取 `readColorNode` 所需的源数据，供 OOXML 公共解析使用。 */
function readColorNode(node: Element | null | undefined, theme: OfficeTheme) {
  if (!node) return undefined;
  const kind = localName(node);
  const base =
    kind === 'srgbclr'
      ? normalizeHex(attr(node, 'val'))
      : kind === 'schemeclr'
      ? normalizeHex(resolveOfficeThemeColor(attr(node, 'val'), theme))
      : kind === 'sysclr'
      ? normalizeHex(attr(node, 'lastClr') ?? attr(node, 'val'))
      : kind === 'prstclr'
      ? normalizeHex(attr(node, 'val'))
      : normalizeHex(
          readColorNode(
            childByLocalName(node, 'srgbClr') ??
              childByLocalName(node, 'schemeClr') ??
              childByLocalName(node, 'sysClr') ??
              childByLocalName(node, 'prstClr'),
            theme,
          ),
        );
  if (!base) return undefined;

  const transforms = Array.from(node.children)
    .map((child) => ({
      type: localName(child),
      val: Number(attr(child, 'val') ?? 0),
    }))
    .filter((item) =>
      [
        'tint',
        'shade',
        'lummod',
        'lumoff',
        'huemod',
        'hueoff',
        'satmod',
        'satoff',
        'alpha',
      ].includes(item.type),
    );

  let alpha = 1;
  let { r, g, b } = hexToRgb(base);
  let hsl = rgbToHsl(r, g, b);

  transforms.forEach((transform) => {
    const ratio = transform.val / 100000;
    switch (transform.type) {
      case 'tint':
        hsl.l = clamp01(hsl.l + (1 - hsl.l) * ratio);
        break;
      case 'shade':
      case 'lummod':
        hsl.l = clamp01(hsl.l * ratio);
        break;
      case 'lumoff':
        hsl.l = clamp01(hsl.l + ratio);
        break;
      case 'huemod':
        hsl.h *= ratio;
        break;
      case 'hueoff':
        hsl.h += transform.val / 60000;
        break;
      case 'satmod':
        hsl.s = clamp01(hsl.s * ratio);
        break;
      case 'satoff':
        hsl.s = clamp01(hsl.s + ratio);
        break;
      case 'alpha':
        alpha = clamp01(ratio);
        break;
      default:
        break;
    }
  });

  const rgb = hslToRgb(hsl.h, hsl.s, hsl.l);
  return alpha < 1
    ? `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${alpha})`
    : rgbToHex(rgb.r, rgb.g, rgb.b);
}

/** 读取 `readFillValue` 所需的源数据，供 OOXML 公共解析使用。 */
function readFillValue(
  node: Element | null | undefined,
  theme: OfficeTheme,
): OfficeChartColor | undefined {
  if (!node) return undefined;
  const kind = localName(node);
  if (kind === 'gradfill') {
    const stops = childrenByLocalName(childByLocalName(node, 'gsLst'), 'gs')
      .map((stop) => {
        const offset = clamp01(Number(attr(stop, 'pos') ?? 0) / 100000);
        const color =
          readColorNode(
            childByLocalName(stop, 'srgbClr') ??
              childByLocalName(stop, 'schemeClr') ??
              childByLocalName(stop, 'sysClr') ??
              childByLocalName(stop, 'prstClr'),
            theme,
          ) ?? undefined;
        return color ? { offset, color } : undefined;
      })
      .filter((stop): stop is OfficeChartColorStop => Boolean(stop))
      .sort((a, b) => a.offset - b.offset);

    if (!stops.length) return undefined;

    const angle =
      Number(attr(childByLocalName(node, 'lin'), 'ang') ?? 5400000) / 60000;
    const radians = (angle * Math.PI) / 180;
    const x = 0.5 - Math.cos(radians) / 2;
    const y = 0.5 - Math.sin(radians) / 2;
    const x2 = 0.5 + Math.cos(radians) / 2;
    const y2 = 0.5 + Math.sin(radians) / 2;

    return {
      type: 'linear',
      x,
      y,
      x2,
      y2,
      global: false,
      colorStops: stops,
    };
  }

  return readColorNode(node, theme);
}

/** 读取 `readLineWidth` 所需的源数据，供 OOXML 公共解析使用。 */
function readLineWidth(node: Element | null | undefined) {
  const width = Number(attr(node, 'w'));
  if (!Number.isFinite(width) || width <= 0) return undefined;
  return width / 9525;
}

/** 读取 `readPositiveNumber` 所需的源数据，供 OOXML 公共解析使用。 */
function readPositiveNumber(node: Element | null | undefined) {
  const value = Number(attr(node, 'val'));
  return Number.isFinite(value) && value > 0 ? value : undefined;
}

/** 读取 `readOfPieSecondPlotCount` 所需的源数据，供 OOXML 公共解析使用。 */
function readOfPieSecondPlotCount(chartNode: Element, pointCount: number) {
  if (pointCount <= 1) return 0;
  const splitPos = readPositiveNumber(childByLocalName(chartNode, 'splitPos'));
  const splitType = attr(childByLocalName(chartNode, 'splitType'), 'val');
  if ((splitType === 'pos' || !splitType) && splitPos) {
    return Math.max(1, Math.min(pointCount - 1, Math.round(splitPos)));
  }

  // Some WPS/Office files omit splitType/splitPos after saving. In that case
  // the extra dPt after the real data points styles the aggregate "Other"
  // slice, while the last two real points are expanded in the secondary plot.
  return Math.max(1, Math.min(pointCount - 1, 2));
}

/** 判断 `isPieLikeChart` 对应的条件是否成立。 */
function isPieLikeChart(type: OfficeChartType, ofPieType?: 'bar' | 'pie') {
  return type === 'pie' || type === 'doughnut' || Boolean(ofPieType);
}

/** 读取 `readShowDataLabels` 所需的源数据，供 OOXML 公共解析使用。 */
function readShowDataLabels(chartNode: Element | null) {
  return descendantsByLocalName(chartNode, 'dLbls').some((labelsNode) => {
    const showVal = childByLocalName(labelsNode, 'showVal');
    return attr(showVal, 'val') === '1' || attr(showVal, 'val') === 'true';
  });
}

/** 读取 `readDataLabels` 所需的源数据，供 OOXML 公共解析使用。 */
function readDataLabels(
  labelsNode: Element | null | undefined,
): OfficeDataLabels | undefined {
  if (!labelsNode) return undefined;
  const dataLabels: OfficeDataLabels = {};
  const readBool = (name: string) => {
    const value = attr(childByLocalName(labelsNode, name), 'val');
    if (value === undefined) return undefined;
    return value === '1' || value === 'true';
  };
  const deleted = attr(childByLocalName(labelsNode, 'delete'), 'val');
  const position = attr(childByLocalName(labelsNode, 'dLblPos'), 'val');
  const separator = textContent(
    childByLocalName(labelsNode, 'separator'),
  ).trim();

  if (deleted === '1' || deleted === 'true') dataLabels.delete = true;
  if (position) dataLabels.position = position;
  if (separator) dataLabels.separator = separator;
  const showLegendKey = readBool('showLegendKey');
  const showVal = readBool('showVal');
  const showCatName = readBool('showCatName');
  const showSerName = readBool('showSerName');
  const showPercent = readBool('showPercent');
  const showBubbleSize = readBool('showBubbleSize');
  const showLeaderLines = readBool('showLeaderLines');
  if (showLegendKey !== undefined) dataLabels.showLegendKey = showLegendKey;
  if (showVal !== undefined) dataLabels.showVal = showVal;
  if (showCatName !== undefined) dataLabels.showCatName = showCatName;
  if (showSerName !== undefined) dataLabels.showSerName = showSerName;
  if (showPercent !== undefined) dataLabels.showPercent = showPercent;
  if (showBubbleSize !== undefined) dataLabels.showBubbleSize = showBubbleSize;
  if (showLeaderLines !== undefined)
    dataLabels.showLeaderLines = showLeaderLines;
  return Object.keys(dataLabels).length ? dataLabels : undefined;
}

/** 读取 `readLegendPosition` 所需的源数据，供 OOXML 公共解析使用。 */
function readLegendPosition(chartNode: Element | null) {
  const value = attr(childByLocalName(chartNode, 'legendPos'), 'val');
  if (value === 'b') return 'bottom';
  if (value === 'l') return 'left';
  if (value === 'r') return 'right';
  return 'top';
}

/** 读取 `readFontFamily` 所需的源数据，供 OOXML 公共解析使用。 */
function readFontFamily(node: Element | null | undefined) {
  const latin = attr(childByLocalName(node, 'latin'), 'typeface');
  const eastAsia = attr(childByLocalName(node, 'ea'), 'typeface');
  const complex = attr(childByLocalName(node, 'cs'), 'typeface');
  const value = [eastAsia, latin, complex]
    .filter((item) => item && !item.startsWith('+'))
    .join(', ');
  return value || undefined;
}

/** 读取 `readLegendVisible` 所需的源数据，供 OOXML 公共解析使用。 */
function readLegendVisible(chartNode: Element | null) {
  if (!chartNode) return false;
  const deleted = attr(childByLocalName(chartNode, 'delete'), 'val');
  return deleted !== '1' && deleted !== 'true';
}

/** 读取 `readLegendStyle` 所需的源数据，供 OOXML 公共解析使用。 */
function readLegendStyle(
  chartNode: Element | null,
  theme: OfficeTheme,
): OfficeChartModel['legendStyle'] | undefined {
  const runProps = descendantByLocalName(chartNode, 'defRPr');
  const fontSize = Number(attr(runProps, 'sz'));
  const color = readFillColor(runProps, theme);
  const fontFamily = readFontFamily(runProps);
  const bold = attr(runProps, 'b');
  const italic = attr(runProps, 'i');
  const textStyle = {
    color,
    fontFamily,
    fontSize:
      Number.isFinite(fontSize) && fontSize > 0 ? fontSize / 100 : undefined,
    fontWeight: bold === '1' || bold === 'true' ? 600 : undefined,
    fontStyle: italic === '1' || italic === 'true' ? 'italic' : undefined,
  };
  const normalizedTextStyle = Object.fromEntries(
    Object.entries(textStyle).filter(([, value]) => value !== undefined),
  );
  return Object.keys(normalizedTextStyle).length
    ? { textStyle: normalizedTextStyle }
    : undefined;
}

/** 读取 `readSeriesMarker` 所需的源数据，供 OOXML 公共解析使用。 */
function readSeriesMarker(seriesNode: Element) {
  const markerNode = childByLocalName(seriesNode, 'marker');
  const symbol = attr(childByLocalName(markerNode, 'symbol'), 'val');
  const size = Number(attr(childByLocalName(markerNode, 'size'), 'val'));
  return {
    symbol: symbol ? symbol.toLowerCase() : undefined,
    size: Number.isFinite(size) && size > 0 ? size : undefined,
  };
}

/** 执行 `looksLikeDateFormat` 封装的 OOXML 公共解析处理步骤。 */
function looksLikeDateFormat(formatCode: string) {
  return /[ymdhs]/i.test(formatCode) && !/^general$/i.test(formatCode);
}

/** 把输入格式化为 `formatDateFromSerial` 返回的展示值。 */
function formatDateFromSerial(
  serial: number,
  formatCode: string,
  date1904: boolean,
) {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  const date = new Date(epoch + serial * 24 * 60 * 60 * 1000);
  const year = String(date.getUTCFullYear());
  const month = String(date.getUTCMonth() + 1);
  const paddedMonth = month.padStart(2, '0');
  const day = String(date.getUTCDate());
  const paddedDay = day.padStart(2, '0');

  return formatCode
    .replace(/yyyy/g, year)
    .replace(/yy/g, year.slice(-2))
    .replace(/mm/g, paddedMonth)
    .replace(/m/g, month)
    .replace(/dd/g, paddedDay)
    .replace(/d/g, day);
}

/** 把输入格式化为 `formatCacheValue` 返回的展示值。 */
function formatCacheValue(
  value: string,
  formatCode: string,
  date1904: boolean,
) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return value;
  }

  if (!formatCode || !looksLikeDateFormat(formatCode)) {
    return value;
  }

  return formatDateFromSerial(numeric, formatCode, date1904);
}

/** 读取 `readDate1904` 所需的源数据，供 OOXML 公共解析使用。 */
function readDate1904(chartSpace: Element | null) {
  const date1904 = childByLocalName(chartSpace, 'date1904');
  return attr(date1904, 'val') === '1' || attr(date1904, 'val') === 'true';
}

/** 查找 `findChartNodes` 对应的目标数据。 */
function findChartNodes(plotArea: Element | null) {
  return Array.from(plotArea?.children ?? []).filter((child) =>
    (child.localName.split(':').pop() ?? child.localName)
      .toLowerCase()
      .endsWith('chart'),
  );
}

/** 读取 `readChartPlot` 所需的源数据，供 OOXML 公共解析使用。 */
function readChartPlot(
  chartNode: Element,
  theme: OfficeTheme,
  date1904: boolean,
) {
  const type = normalizeType(chartNode);
  const chartKind = localName(chartNode);
  const seriesNodes = childrenByLocalName(chartNode, 'ser');
  const firstSeries = seriesNodes[0];
  const plotDataLabels = readDataLabels(
    childrenByLocalName(chartNode, 'dLbls')[0],
  );
  const grouping = attr(childByLocalName(chartNode, 'grouping'), 'val');
  const stacking: OfficeChartSeries['stacking'] =
    grouping === 'stacked' || grouping === 'percentStacked'
      ? grouping
      : undefined;
  const stackGroup = stacking ? `office-chart-${type}` : undefined;
  const categories = readCacheValues(
    descendantByLocalName(firstSeries, 'cat'),
    date1904,
  ).map(decodeMojibake);
  const firstSliceAngle = Number(
    attr(childByLocalName(chartNode, 'firstSliceAng'), 'val'),
  );
  const gapWidth = readPositiveNumber(childByLocalName(chartNode, 'gapWidth'));
  const overlapValue = Number(
    attr(childByLocalName(chartNode, 'overlap'), 'val'),
  );
  const overlap = Number.isFinite(overlapValue) ? overlapValue : undefined;
  const series = seriesNodes.map((seriesNode, index) => ({
    name:
      firstText(descendantByLocalName(seriesNode, 'tx')) ||
      `Series ${index + 1}`,
    type,
    stacking,
    stackGroup,
    gapWidth,
    overlap,
    values: readNumericValues(descendantByLocalName(seriesNode, 'val')),
    color: readSeriesColorWithTheme(seriesNode, theme),
    lineWidth: readLineWidth(
      childByLocalName(childByLocalName(seriesNode, 'spPr'), 'ln'),
    ),
    pointColors: readPointColors(seriesNode, theme),
    pointStyles: readPointStyles(seriesNode, theme),
    dataLabels:
      readDataLabels(childrenByLocalName(seriesNode, 'dLbls')[0]) ??
      plotDataLabels,
    smooth:
      attr(childByLocalName(seriesNode, 'smooth'), 'val') === '1' ||
      attr(childByLocalName(seriesNode, 'smooth'), 'val') === 'true',
    marker: readSeriesMarker(seriesNode),
  }));
  const firstValueCount = series[0]?.values.length ?? 0;
  const ofPieTypeValue = attr(childByLocalName(chartNode, 'ofPieType'), 'val');
  const ofPieType =
    chartKind === 'ofpiechart' &&
    (ofPieTypeValue === 'bar' || ofPieTypeValue === 'pie')
      ? ofPieTypeValue
      : undefined;

  return {
    type,
    categories,
    series,
    dataLabels: plotDataLabels,
    holeSize:
      type === 'doughnut'
        ? Number(attr(childByLocalName(chartNode, 'holeSize'), 'val') ?? 0) ||
          undefined
        : undefined,
    startAngle: isPieLikeChart(type, ofPieType)
      ? Number.isFinite(firstSliceAngle)
        ? firstSliceAngle
        : 0
      : undefined,
    ofPieType,
    ofPieSecondPlotCount: ofPieType
      ? readOfPieSecondPlotCount(chartNode, firstValueCount)
      : undefined,
    secondPieSize: ofPieType
      ? readPositiveNumber(childByLocalName(chartNode, 'secondPieSize'))
      : undefined,
    gapWidth,
    overlap,
    radarStyle:
      type === 'radar'
        ? attr(childByLocalName(chartNode, 'radarStyle'), 'val')
        : undefined,
  };
}

/** 执行 `niceRadarMax` 封装的 OOXML 公共解析处理步骤。 */
function niceRadarMax(value: number) {
  if (!Number.isFinite(value) || value <= 0) return 1;
  const magnitude = 10 ** Math.floor(Math.log10(value));
  const normalized = value / magnitude;
  const step = normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  return step * magnitude;
}

/** 根据输入构建 `buildRadarIndicators` 返回的标准化结果。 */
function buildRadarIndicators(
  categories: string[],
  series: OfficeChartSeries[],
) {
  return categories.map((name, index) => {
    const maxValue = series.reduce(
      (acc, item) => Math.max(acc, item.values[index] ?? 0),
      0,
    );
    return {
      name,
      max: niceRadarMax(maxValue),
    };
  });
}

/** 解析 `parseOfficeChartXml` 接收的数据，并返回 OOXML 公共解析结果。 */
export function parseOfficeChartXml(
  xml: string,
  theme: OfficeTheme = DEFAULT_OFFICE_THEME,
): OfficeChartModel {
  // 先把 OOXML chart 统一成中间模型，组件层不直接依赖复杂的 c:* XML 结构。
  const doc = parseXml(xml);
  const chartSpace = doc.documentElement;
  const chart = descendantByLocalName(chartSpace, 'chart');
  const plotArea = descendantByLocalName(chart, 'plotArea');
  const date1904 = readDate1904(chartSpace);
  const plots = findChartNodes(plotArea).map((chartNode) =>
    readChartPlot(chartNode, theme, date1904),
  );
  const primaryPlot = plots[0];
  const categories = primaryPlot?.categories ?? [];
  const series = plots.flatMap((plot) => plot.series);
  const type = primaryPlot?.type ?? 'unknown';
  const rawTitle = firstText(childByLocalName(chart, 'title'));
  const autoTitleDeleted = attr(
    childByLocalName(chart, 'autoTitleDeleted'),
    'val',
  );
  const title =
    rawTitle ||
    (childByLocalName(chart, 'title') &&
    autoTitleDeleted !== '1' &&
    autoTitleDeleted !== 'true'
      ? series[0]?.name
      : undefined);

  return {
    type,
    title: title || undefined,
    categories,
    series,
    showLegend: readLegendVisible(childByLocalName(chart, 'legend')),
    legendPosition: readLegendPosition(childByLocalName(chart, 'legend')),
    legendStyle: readLegendStyle(childByLocalName(chart, 'legend'), theme),
    showDataLabels: readShowDataLabels(plotArea),
    radarIndicators:
      type === 'radar' ? buildRadarIndicators(categories, series) : undefined,
    holeSize: primaryPlot?.holeSize,
    startAngle: primaryPlot?.startAngle,
    ofPieType:
      primaryPlot?.ofPieType === 'bar' || primaryPlot?.ofPieType === 'pie'
        ? primaryPlot.ofPieType
        : undefined,
    ofPieSecondPlotCount: primaryPlot?.ofPieSecondPlotCount,
    secondPieSize: primaryPlot?.secondPieSize,
    gapWidth: primaryPlot?.gapWidth,
    overlap: primaryPlot?.overlap,
    radarStyle: primaryPlot?.radarStyle,
  };
}

/** 解析并确定 `resolveSeriesColor` 对应的引用或配置。 */
function resolveSeriesColor(series: OfficeChartSeries, index: number) {
  return series.color ?? DEFAULT_COLORS[index % DEFAULT_COLORS.length];
}

/** 解析并确定 `resolveCategories` 对应的引用或配置。 */
function resolveCategories(chart: OfficeChartModel) {
  if (chart.categories.length) {
    return chart.categories;
  }

  const maxLength = Math.max(
    ...chart.series.map((series) => series.values.length),
    0,
  );
  return Array.from({ length: maxLength }, (_, index) => String(index + 1));
}

/** 将输入标准化为 `normalizeSeriesType` 返回的结构。 */
function normalizeSeriesType(type: OfficeChartType) {
  if (type === 'column' || type === 'bar') return 'bar';
  if (type === 'area') return 'line';
  if (type === 'scatter' || type === 'bubble') return 'scatter';
  if (type === 'radar') return 'radar';
  if (type === 'pie' || type === 'doughnut') return 'pie';
  return 'line';
}

/** 执行 `sanitizeMapRegionName` 封装的 OOXML 公共解析处理步骤。 */
function sanitizeMapRegionName(name: string) {
  return name
    .replace(
      /特别行政区$|壮族自治区$|回族自治区$|维吾尔自治区$|自治区$|省$|市$/g,
      '',
    )
    .trim();
}

/** 执行 `scaleRoseRadius` 封装的 OOXML 公共解析处理步骤。 */
function scaleRoseRadius(
  radius: [string, string] | undefined,
): [string, string] | undefined {
  if (!radius) return undefined;
  const inner = Number(radius[0].replace(/%$/, ''));
  const outer = Number(radius[1].replace(/%$/, ''));
  if (!Number.isFinite(inner) || !Number.isFinite(outer) || outer <= 0)
    return radius;
  const fittedOuter = Math.min(58, outer);
  const fittedInner = Math.max(
    0,
    Math.min(fittedOuter - 4, Math.round((inner / outer) * fittedOuter)),
  );
  return [`${fittedInner}%`, `${fittedOuter}%`];
}

/** 根据输入构建 `buildLegend` 返回的标准化结果。 */
function buildLegend(chart: OfficeChartModel, itemCount = chart.series.length) {
  if (chart.showLegend === false || itemCount <= 0) return undefined;

  const base = {
    type: 'scroll' as const,
    itemWidth: chart.legendStyle?.itemWidth ?? 10,
    itemHeight: chart.legendStyle?.itemHeight ?? 10,
    textStyle: {
      ...OFFICE_TEXT_STYLE,
      ...chart.legendStyle?.textStyle,
    },
  } as const;

  switch (chart.legendPosition) {
    case 'bottom':
      return {
        ...base,
        bottom: 4,
      };
    case 'left':
      return {
        ...base,
        left: 8,
        top: chart.title ? 32 : 8,
        orient: 'vertical' as const,
      };
    case 'right':
      return {
        ...base,
        right: 8,
        top: chart.title ? 32 : 8,
        orient: 'vertical' as const,
      };
    default:
      return {
        ...base,
        top: chart.title ? 30 : 8,
      };
  }
}

/** 根据输入构建 `buildChartGrid` 返回的标准化结果。 */
function buildChartGrid(chart: OfficeChartModel) {
  const isBottomLegend = chart.legendPosition === 'bottom';
  const isSideLegend =
    chart.legendPosition === 'left' || chart.legendPosition === 'right';
  return {
    left: isSideLegend ? 70 : 40,
    right: isSideLegend ? 70 : 24,
    top: chart.title ? 56 : chart.legendPosition === 'top' ? 40 : 24,
    bottom: isBottomLegend ? 56 : 32,
    containLabel: true,
  };
}

/** 根据输入构建 `buildOfficeTitle` 返回的标准化结果。 */
function buildOfficeTitle(chart: OfficeChartModel) {
  return chart.title
    ? {
        text: chart.title,
        left: 'center',
        top: 8,
        textStyle: {
          fontSize: 14,
          fontWeight: 600,
          color: '#111827',
          fontFamily: OFFICE_FONT_FAMILY,
        },
      }
    : undefined;
}

/** 解析并确定 `resolveOfficePieStartAngle` 对应的引用或配置。 */
function resolveOfficePieStartAngle(chart: OfficeChartModel) {
  const officeAngle = chart.startAngle ?? 0;
  return (((90 - officeAngle) % 360) + 360) % 360;
}

/** 解析并确定 `resolveOfficeRadarStartAngle` 对应的引用或配置。 */
function resolveOfficeRadarStartAngle(chart: OfficeChartModel) {
  return chart.radarStartAngle ?? 90;
}

/** 解析并确定 `resolveOfficeRadarRadius` 对应的引用或配置。 */
function resolveOfficeRadarRadius(chart: OfficeChartModel) {
  return (
    chart.radarRadius ?? (chart.legendPosition === 'bottom' ? '62%' : '68%')
  );
}

/** 执行 `reorderRadarAxes` 封装的 OOXML 公共解析处理步骤。 */
function reorderRadarAxes<T>(items: T[]) {
  if (items.length <= 2) return items.slice();
  return [items[0], ...items.slice(1).reverse()];
}

/** 解析并确定 `resolveOfficeRadarCenter` 对应的引用或配置。 */
function resolveOfficeRadarCenter(chart: OfficeChartModel): [string, string] {
  if (chart.legendPosition === 'bottom')
    return ['50%', chart.title ? '52%' : '48%'];
  if (chart.legendPosition === 'top')
    return ['50%', chart.title ? '58%' : '56%'];
  return ['50%', chart.title ? '56%' : '52%'];
}

/** 解析并确定 `resolveBarWidthFromGap` 对应的引用或配置。 */
function resolveBarWidthFromGap(
  gapWidth: number | undefined,
  seriesCount: number,
  overlap?: number,
) {
  if (!Number.isFinite(gapWidth) || gapWidth === undefined) return undefined;
  const visibleSeriesCount = Math.max(1, seriesCount);
  const overlapRatio = Math.max(-1, Math.min(1, (overlap ?? 0) / 100));
  const effectiveSeriesCount = Math.max(
    1,
    visibleSeriesCount - Math.max(0, overlapRatio) * (visibleSeriesCount - 1),
  );
  const categoryWidth = 72;
  const width = categoryWidth / (effectiveSeriesCount + gapWidth / 100);
  return Math.max(6, Math.min(46, Math.round(width)));
}

/** 读取 `readPieLabelPosition` 所需的源数据，供 OOXML 公共解析使用。 */
function readPieLabelPosition(
  labels?: OfficeDataLabels,
  fallback: 'outside' | 'inside' = 'outside',
) {
  const position = labels?.position?.toLowerCase();
  if (!position) return fallback;
  if (position.includes('in')) return 'inside';
  if (position.includes('out')) return 'outside';
  return fallback;
}

/** 读取 `readCartesianLabelPosition` 所需的源数据，供 OOXML 公共解析使用。 */
function readCartesianLabelPosition(
  labels?: OfficeDataLabels,
  horizontal = false,
) {
  const position = labels?.position?.toLowerCase();
  if (!position) return horizontal ? 'right' : 'top';
  if (position === 'ctr' || position === 'center') return 'inside';
  if (position === 'inbase') return horizontal ? 'insideLeft' : 'insideBottom';
  if (position === 'inend') return horizontal ? 'insideRight' : 'insideTop';
  if (position === 'outend') return horizontal ? 'right' : 'top';
  if (position.includes('base')) return horizontal ? 'left' : 'bottom';
  if (position.includes('end')) return horizontal ? 'right' : 'top';
  return horizontal ? 'right' : 'top';
}

/** 根据输入构建 `buildDataLabelFormatter` 返回的标准化结果。 */
function buildDataLabelFormatter(
  labels: OfficeDataLabels | undefined,
  categories: string[],
) {
  const showValue = labels?.showVal ?? false;
  const showCategory = labels?.showCatName ?? false;
  const showSeries = labels?.showSerName ?? false;
  const showPercent = labels?.showPercent ?? false;
  const separator = labels?.separator ?? '\n';
  return (params: unknown) => {
    const item = params as {
      /** 图表转换配置的可读名称。 */
      name?: string;
      /** 当前结构 保存的解析值或业务值。 */
      value?: unknown;
      /** 当前结构 在所属集合中的位置索引。 */
      dataIndex?: number;
      /** 面向界面展示的解析完成百分比。 */
      percent?: number;
      /** 图表转换配置的 seriesName 文本值。 */
      seriesName?: string;
    };
    const value = Array.isArray(item.value)
      ? item.value[item.value.length - 1]
      : item.value;
    const category =
      item.name ??
      (item.dataIndex !== undefined ? categories[item.dataIndex] : undefined);
    const parts: string[] = [];
    if (showSeries && item.seriesName) parts.push(item.seriesName);
    if (showCategory && category) parts.push(category);
    if (showValue && value !== undefined) parts.push(String(value));
    if (showPercent && item.percent !== undefined)
      parts.push(`${item.percent}%`);
    return parts.join(separator).trim();
  };
}

/** 执行 `shouldShowDataLabels` 封装的 OOXML 公共解析处理步骤。 */
function shouldShowDataLabels(
  labels: OfficeDataLabels | undefined,
  chartShowDataLabels?: boolean,
) {
  if (labels?.delete) return false;
  if (!labels) return Boolean(chartShowDataLabels);
  const explicitFlags = [
    labels.showVal,
    labels.showCatName,
    labels.showSerName,
    labels.showPercent,
  ].filter((value) => value !== undefined);
  if (explicitFlags.length) return explicitFlags.some(Boolean);
  return Boolean(chartShowDataLabels);
}

/** 根据输入构建 `buildCartesianDataLabelConfig` 返回的标准化结果。 */
function buildCartesianDataLabelConfig(
  labels: OfficeDataLabels | undefined,
  chartShowDataLabels: boolean | undefined,
  categories: string[],
  horizontal = false,
) {
  const effectiveLabels =
    labels ?? (chartShowDataLabels ? { showVal: true } : undefined);
  return {
    show: shouldShowDataLabels(effectiveLabels, chartShowDataLabels),
    position: readCartesianLabelPosition(effectiveLabels, horizontal),
    formatter: buildDataLabelFormatter(effectiveLabels, categories),
    color: '#334155',
    fontFamily: OFFICE_FONT_FAMILY,
  };
}

/** 根据输入构建 `buildPieDataLabelConfig` 返回的标准化结果。 */
function buildPieDataLabelConfig(
  labels: OfficeDataLabels | undefined,
  showDataLabels?: boolean,
) {
  const position = readPieLabelPosition(labels, 'outside');
  const showValue = labels?.showVal ?? showDataLabels;
  const showCategory = labels?.showCatName ?? false;
  const showSeries = labels?.showSerName ?? false;
  const showPercent = labels?.showPercent ?? false;
  const separator = labels?.separator ?? '\n';
  const formatter = (params: unknown) => {
    const item = params as {
      /** 图表转换配置的可读名称。 */
      name?: string;
      /** 当前结构 当前步骤需要处理的原始或标准化数据。 */
      data?: {
        /** 图表转换配置的可读名称。 */ name?: string;
      };
      /** 当前结构 保存的解析值或业务值。 */
      value?: number;
      /** 面向界面展示的解析完成百分比。 */
      percent?: number;
      /** 图表转换配置的 seriesName 文本值。 */
      seriesName?: string;
    };
    const parts: string[] = [];
    if (showSeries && item.seriesName) parts.push(item.seriesName);
    if (showCategory && (item.data?.name ?? item.name))
      parts.push(item.data?.name ?? item.name ?? '');
    if (showValue && item.value !== undefined) parts.push(String(item.value));
    if (showPercent && item.percent !== undefined)
      parts.push(`${item.percent}%`);
    return parts.join(separator).trim();
  };

  return {
    show:
      !labels?.delete &&
      Boolean(
        showValue ||
          showCategory ||
          showSeries ||
          showPercent ||
          showDataLabels,
      ),
    position,
    formatter,
  };
}

/** 根据输入构建 `buildOfPieChartOption` 返回的标准化结果。 */
function buildOfPieChartOption(
  chart: OfficeChartModel,
  categories: string[],
  palette: string[],
): EChartsOption {
  const sourceSeries = chart.series[0];
  const pieLabels = chart.dataLabels ?? sourceSeries?.dataLabels;
  const values = sourceSeries?.values ?? [];
  const secondCount = Math.max(
    1,
    Math.min(
      values.length - 1,
      chart.ofPieSecondPlotCount ?? Math.ceil(values.length / 3),
    ),
  );
  const splitIndex = Math.max(1, values.length - secondCount);
  const mainNames = categories.slice(0, splitIndex);
  const secondaryNames = categories.slice(splitIndex, values.length);
  const secondaryValues = values.slice(splitIndex);
  const secondaryTotal = secondaryValues.reduce((sum, value) => sum + value, 0);
  const otherName = '其他';
  const otherStyle = buildPieItemStyle(
    sourceSeries,
    categories.length,
    palette,
  );
  const startAngle = resolveOfficePieStartAngle(chart);
  const total = values.reduce((sum, value) => sum + value, 0);
  const beforeOtherTotal = values
    .slice(0, splitIndex)
    .reduce((sum, value) => sum + value, 0);
  const otherStart = total
    ? startAngle - (beforeOtherTotal / total) * 360
    : startAngle;
  const otherEnd = total
    ? otherStart - (secondaryTotal / total) * 360
    : startAngle;
  const otherMid = ((otherStart + otherEnd) / 2) * (Math.PI / 180);
  const connectorAnchorX = 34 + Math.cos(otherMid) * 23;
  const connectorAnchorY = (chart.title ? 58 : 52) - Math.sin(otherMid) * 23;
  const mainData = [
    ...mainNames.map((name, index) => ({
      name,
      value: values[index] ?? 0,
      itemStyle: buildPieItemStyle(sourceSeries, index, palette),
    })),
    {
      name: otherName,
      value: secondaryTotal,
      itemStyle: otherStyle,
      tooltip: {
        formatter: `${otherName}<br/>${
          sourceSeries?.name ?? ''
        }: ${secondaryTotal}`,
      },
    },
  ];
  const secondarySize = Math.max(
    28,
    Math.min(70, Math.round(52 * ((chart.secondPieSize ?? 75) / 75))),
  );
  const legend = buildLegend(chart, categories.length) as
    | Record<string, unknown>
    | undefined;
  const tooltip = {
    trigger: 'item' as const,
    confine: true,
    appendToBody: true,
    backgroundColor: 'rgba(15, 23, 42, 0.96)',
    borderColor: 'rgba(15, 23, 42, 0.96)',
    textStyle: {
      color: '#fff',
      fontFamily: OFFICE_FONT_FAMILY,
    },
    formatter: (params: unknown) => {
      const item = params as {
        /** 图表转换配置的 componentSubType 文本值。 */
        componentSubType?: string;
        /** 当前结构 当前步骤需要处理的原始或标准化数据。 */
        data?: {
          /** 图表转换配置的可读名称。 */ name?: string;
        };
        /** 图表转换配置的可读名称。 */
        name?: string;
        /** 图表转换配置的 seriesName 文本值。 */
        seriesName?: string;
        /** 当前结构 保存的解析值或业务值。 */
        value?: unknown;
      };
      const value =
        typeof item.value === 'number'
          ? item.value
          : Array.isArray(item.value)
          ? item.value[0]
          : '';
      const name =
        item.componentSubType === 'bar'
          ? item.seriesName
          : item.data?.name ?? item.name ?? '';
      return `${name}<br/>${sourceSeries?.name ?? ''}: ${value}`;
    },
  };

  const secondarySeries =
    chart.ofPieType === 'pie'
      ? [
          {
            type: 'pie' as const,
            radius: ['0%', `${secondarySize}%`] as [string, string],
            center: ['72%', chart.title ? '58%' : '52%'] as [string, string],
            startAngle,
            avoidLabelOverlap: true,
            label: buildPieDataLabelConfig(pieLabels, chart.showDataLabels),
            labelLayout: {
              hideOverlap: true,
            },
            labelLine: {
              length: 12,
              length2: 8,
              smooth: true,
            },
            data: secondaryNames.map((name, index) => ({
              name,
              value: secondaryValues[index] ?? 0,
              itemStyle: buildPieItemStyle(
                sourceSeries,
                splitIndex + index,
                palette,
              ),
            })),
          },
        ]
      : secondaryNames.map((name, index) => ({
          name,
          type: 'bar' as const,
          stack: 'office-of-pie-secondary',
          barWidth: Math.max(
            26,
            Math.min(46, Math.round(secondarySize * 0.72)),
          ),
          xAxisIndex: 0,
          yAxisIndex: 0,
          data: [
            {
              value: secondaryValues[index] ?? 0,
              itemStyle: buildPieItemStyle(
                sourceSeries,
                splitIndex + index,
                palette,
              ),
            },
          ],
          label: {
            show: chart.showDataLabels,
            position: 'inside' as const,
            color: '#334155',
            fontFamily: OFFICE_FONT_FAMILY,
          },
          emphasis: {
            itemStyle: {
              shadowBlur: 6,
              shadowColor: 'rgba(15, 23, 42, 0.18)',
            },
          },
        }));

  return {
    animation: false,
    backgroundColor: '#fff',
    color: palette,
    textStyle: OFFICE_TEXT_STYLE,
    title: buildOfficeTitle(chart),
    tooltip,
    legend: legend
      ? {
          ...legend,
          data: categories,
        }
      : undefined,
    grid:
      chart.ofPieType === 'bar'
        ? {
            left: '64%',
            top: chart.title ? '34%' : '26%',
            width: '16%',
            height: chart.title ? '46%' : '52%',
            containLabel: false,
          }
        : undefined,
    xAxis:
      chart.ofPieType === 'bar'
        ? {
            type: 'category' as const,
            data: [''],
            show: false,
          }
        : undefined,
    yAxis:
      chart.ofPieType === 'bar'
        ? {
            type: 'value' as const,
            min: 0,
            max: secondaryTotal || undefined,
            show: false,
          }
        : undefined,
    graphic:
      chart.ofPieType === 'bar'
        ? [
            {
              type: 'line',
              left: `${connectorAnchorX}%`,
              top: `${connectorAnchorY}%`,
              shape: { x1: 0, y1: 0, x2: 132, y2: chart.title ? -28 : -24 },
              style: { stroke: '#cbd5e1', lineWidth: 1 },
              silent: true,
            },
            {
              type: 'line',
              left: `${connectorAnchorX}%`,
              top: `${connectorAnchorY}%`,
              shape: { x1: 0, y1: 0, x2: 132, y2: chart.title ? 44 : 40 },
              style: { stroke: '#cbd5e1', lineWidth: 1 },
              silent: true,
            },
          ]
        : undefined,
    series: [
      {
        type: 'pie' as const,
        radius: ['0%', '46%'] as [string, string],
        center: ['34%', chart.title ? '58%' : '52%'] as [string, string],
        startAngle,
        avoidLabelOverlap: true,
        label: buildPieDataLabelConfig(pieLabels, chart.showDataLabels),
        labelLayout: {
          hideOverlap: true,
        },
        labelLine: {
          length: 12,
          length2: 8,
          smooth: true,
        },
        emphasis: {
          scale: false,
          itemStyle: {
            shadowBlur: 8,
            shadowColor: 'rgba(15, 23, 42, 0.18)',
          },
        },
        data: mainData,
      },
      ...secondarySeries,
    ],
  };
}

/** 根据输入构建 `buildOfficeChartOption` 返回的标准化结果。 */
export function buildOfficeChartOption(chart: OfficeChartModel): EChartsOption {
  // 中间模型在这里映射为 ECharts option，PPTX/DOCX/XLSX 共用同一套图表渲染逻辑。
  const categories = resolveCategories(chart);
  const normalizedSeriesTypes = chart.series.map((item) =>
    normalizeSeriesType(item.type ?? chart.type),
  );
  const uniqueSeriesTypes = new Set(normalizedSeriesTypes);
  const isHorizontalBar = chart.type === 'bar';
  const isPie = chart.type === 'pie' || chart.type === 'doughnut';
  const isRadar = chart.type === 'radar';
  const isScatter =
    normalizedSeriesTypes.length > 0 &&
    normalizedSeriesTypes.every((type) => type === 'scatter');
  const usesMixedSeriesTypes = uniqueSeriesTypes.size > 1;
  const palette = chart.series.map((series, index) =>
    resolveSeriesColor(series, index),
  );
  const hasSeries = chart.series.length > 0;

  if (!hasSeries) {
    return {
      animation: false,
      title: chart.title
        ? {
            text: chart.title,
            left: 'center',
            top: 8,
            textStyle: {
              fontSize: 14,
              fontWeight: 600,
            },
          }
        : undefined,
    };
  }

  const radarIndicators =
    chart.radarIndicators ??
    (isRadar ? buildRadarIndicators(categories, chart.series) : undefined);
  const radarDisplayIndicators =
    isRadar && radarIndicators?.length
      ? reorderRadarAxes(radarIndicators)
      : undefined;
  const radarCategories =
    radarDisplayIndicators?.map((indicator) => indicator.name) ?? categories;

  if (chart.type === 'map') {
    const sourceSeries = chart.series[0];
    const values = sourceSeries?.values ?? [];
    const tierNames = Array.from(
      new Set((sourceSeries?.pointLabels ?? []).filter(Boolean)),
    );
    const tierColors = tierNames
      .map((tier) => {
        const index = sourceSeries?.pointLabels?.indexOf(tier) ?? -1;
        return index >= 0 ? sourceSeries?.pointColors?.[index] : undefined;
      })
      .filter((color): color is string => Boolean(color));
    const data = categories.map((name, index) => ({
      name,
      value: values[index] ?? 0,
      labelName: sanitizeMapRegionName(name),
      tierName: sourceSeries?.pointLabels?.[index],
      itemStyle: {
        areaColor: sourceSeries?.pointColors?.[index] ?? '#e5edf8',
        borderColor: '#ffffff',
        borderWidth: 1,
      },
    }));

    return {
      animation: false,
      backgroundColor: '#ffffff',
      textStyle: OFFICE_TEXT_STYLE,
      title: chart.title
        ? {
            text: chart.title,
            subtext: chart.mapRegion,
            left: 'center',
            top: 8,
            textStyle: {
              fontSize: 14,
              fontWeight: 600,
              color: '#111827',
              fontFamily: OFFICE_FONT_FAMILY,
            },
            subtextStyle: {
              color: '#64748b',
              fontFamily: OFFICE_FONT_FAMILY,
            },
          }
        : undefined,
      tooltip: {
        trigger: 'item',
        confine: true,
        appendToBody: true,
        backgroundColor: 'rgba(15, 23, 42, 0.96)',
        borderColor: 'rgba(15, 23, 42, 0.96)',
        textStyle: {
          color: '#fff',
          fontFamily: OFFICE_FONT_FAMILY,
        },
        formatter: (params: unknown) => {
          const item = params as {
            /** 当前结构 当前步骤需要处理的原始或标准化数据。 */
            data?: {
              /** 图表转换配置的 tierName 文本值。 */ tierName?: string;
            };
            /** 图表转换配置的可读名称。 */
            name?: string;
            /** 当前结构 保存的解析值或业务值。 */
            value?: unknown;
          };
          const value = typeof item.value === 'number' ? item.value : '';
          const tier = item.data?.tierName ? `<br/>${item.data.tierName}` : '';
          return `${item.name ?? ''}<br/>${
            sourceSeries?.name ?? chart.mapSeriesName ?? ''
          }: ${value}${tier}`;
        },
      },
      graphic: tierNames.length
        ? {
            type: 'group',
            left: 12,
            bottom: 12,
            children: tierNames.flatMap((name, index) => [
              {
                type: 'rect',
                shape: { x: 0, y: index * 20, width: 10, height: 10 },
                style: { fill: tierColors[index] ?? '#cbd5e1' },
              },
              {
                type: 'text',
                left: 16,
                top: index * 20 - 2,
                style: {
                  text: name,
                  fill: '#334155',
                  font: `12px ${OFFICE_FONT_FAMILY}`,
                },
              },
            ]),
          }
        : undefined,
      series: [
        {
          name: sourceSeries?.name ?? chart.mapSeriesName,
          type: 'map' as const,
          map: chart.mapName ?? 'china',
          roam: true,
          selectedMode: false,
          layoutCenter: ['50%', chart.title ? '56%' : '52%'],
          layoutSize: tierNames.length ? '88%' : '92%',
          zoom: 1.08,
          itemStyle: {
            areaColor: '#eef3f8',
            borderColor: '#f8fafc',
            borderWidth: 1,
          },
          emphasis: {
            label: {
              show: true,
              color: '#0f172a',
              fontFamily: OFFICE_FONT_FAMILY,
            },
            itemStyle: {
              areaColor: '#f59e0b',
              borderColor: '#ffffff',
              borderWidth: 1.2,
            },
          },
          label: {
            show: chart.showDataLabels ?? true,
            color: '#1f2937',
            fontFamily: OFFICE_FONT_FAMILY,
            fontSize: 9,
            formatter: (params: unknown) => {
              const item = params as {
                /** 当前结构 当前步骤需要处理的原始或标准化数据。 */
                data?: {
                  /** 图表转换配置的 labelName 文本值。 */ labelName?: string;
                };
                /** 图表转换配置的可读名称。 */
                name?: string;
              };
              return (
                item.data?.labelName ?? sanitizeMapRegionName(item.name ?? '')
              );
            },
          },
          data,
        },
      ],
    };
  }

  if (isRadar && radarDisplayIndicators?.length) {
    return {
      animation: false,
      backgroundColor: '#fff',
      color: palette,
      textStyle: OFFICE_TEXT_STYLE,
      title: chart.title
        ? {
            text: chart.title,
            left: 'center',
            top: 8,
            textStyle: {
              fontSize: 14,
              fontWeight: 600,
              color: '#111827',
              fontFamily: OFFICE_FONT_FAMILY,
            },
          }
        : undefined,
      tooltip: {
        trigger: 'item',
        confine: true,
        appendToBody: true,
        backgroundColor: 'rgba(15, 23, 42, 0.96)',
        borderColor: 'rgba(15, 23, 42, 0.96)',
        textStyle: {
          color: '#fff',
          fontFamily: OFFICE_FONT_FAMILY,
        },
      },
      legend: buildLegend(chart),
      radar: {
        center: resolveOfficeRadarCenter(chart),
        radius: resolveOfficeRadarRadius(chart),
        startAngle: resolveOfficeRadarStartAngle(chart),
        indicator: radarDisplayIndicators,
        splitNumber: chart.radarSplitNumber ?? 5,
        axisName: {
          color: '#475569',
          fontFamily: OFFICE_FONT_FAMILY,
        },
        splitArea: {
          areaStyle: {
            color: ['rgba(255,255,255,0)', 'rgba(148,163,184,0.04)'],
          },
        },
        splitLine: {
          lineStyle: {
            color: '#e2e8f0',
          },
        },
        axisLine: {
          lineStyle: {
            color: '#cbd5e1',
          },
        },
      },
      series: chart.series.map((item, index) => {
        const color = resolveSeriesColor(item, index);
        return {
          name: item.name,
          type: 'radar',
          areaStyle:
            chart.radarStyle === 'filled' ? { opacity: 0.18 } : undefined,
          data: [
            {
              value: reorderRadarAxes(
                item.values.slice(0, radarDisplayIndicators.length),
              ),
              name: item.name,
            },
          ],
          symbol: item.marker?.symbol ?? 'circle',
          symbolSize: item.marker?.size ?? 6,
          lineStyle: {
            color,
            width: item.lineWidth ?? 2,
          },
          itemStyle: {
            color,
          },
          label: {
            show: shouldShowDataLabels(
              item.dataLabels ?? chart.dataLabels,
              chart.showDataLabels,
            ),
            formatter: buildDataLabelFormatter(
              item.dataLabels ?? chart.dataLabels,
              radarCategories,
            ),
            color: '#334155',
            fontFamily: OFFICE_FONT_FAMILY,
          },
        };
      }),
    };
  }

  if (isPie && chart.ofPieType && categories.length > 1) {
    return buildOfPieChartOption(chart, categories, palette);
  }

  if (isPie && !usesMixedSeriesTypes) {
    const sourceSeries = chart.series[0];
    const pieLabels = chart.dataLabels ?? sourceSeries?.dataLabels;
    const data = categories.map((name, index) => ({
      name,
      value: sourceSeries?.values[index] ?? 0,
      itemStyle: buildPieItemStyle(sourceSeries, index, palette),
    }));
    const innerRadius =
      chart.type === 'doughnut' && chart.holeSize
        ? `${Math.max(
            8,
            Math.min(90, Math.round(68 * (chart.holeSize / 100))),
          )}%`
        : '0%';
    const radius: [string, string] = chart.roseType
      ? scaleRoseRadius(chart.radius) ?? [innerRadius, '58%']
      : chart.radius ?? [innerRadius, '68%'];

    return {
      animation: false,
      backgroundColor: '#fff',
      color: palette,
      textStyle: OFFICE_TEXT_STYLE,
      title: chart.title
        ? {
            text: chart.title,
            left: 'center',
            top: 8,
            textStyle: {
              fontSize: 14,
              fontWeight: 600,
              color: '#111827',
              fontFamily: OFFICE_FONT_FAMILY,
            },
          }
        : undefined,
      tooltip: {
        trigger: 'item',
        confine: true,
        appendToBody: true,
        backgroundColor: 'rgba(15, 23, 42, 0.96)',
        borderColor: 'rgba(15, 23, 42, 0.96)',
        textStyle: {
          color: '#fff',
          fontFamily: OFFICE_FONT_FAMILY,
        },
      },
      legend: buildLegend(chart, categories.length),
      series: [
        {
          type: 'pie' as const,
          radius,
          roseType: chart.roseType,
          startAngle: resolveOfficePieStartAngle(chart),
          padAngle: 0,
          center: ['50%', chart.roseType ? '50%' : chart.title ? '58%' : '50%'],
          avoidLabelOverlap: true,
          label: buildPieDataLabelConfig(pieLabels, chart.showDataLabels),
          labelLayout: {
            hideOverlap: true,
          },
          emphasis: {
            scale: false,
            itemStyle: {
              borderColor: '#ffffff',
              borderWidth: 1,
              shadowBlur: 8,
              shadowColor: 'rgba(15, 23, 42, 0.18)',
            },
          },
          labelLine: {
            length: 12,
            length2: 8,
          },
          data,
        },
      ],
    };
  }

  const series = chart.series.map((item, index) => {
    const seriesType = normalizeSeriesType(item.type ?? chart.type) as
      | 'line'
      | 'bar'
      | 'scatter'
      | 'radar'
      | 'pie';
    const color = resolveSeriesColor(item, index);
    const isBarSeries = seriesType === 'bar';
    const isLineSeries = seriesType === 'line';
    const markerSymbol =
      item.marker?.symbol && item.marker.symbol !== 'none'
        ? item.marker.symbol
        : undefined;
    const hideSymbol = item.marker?.symbol === 'none';
    const isBubbleSeries = item.type === 'bubble';
    const barSeriesCount = chart.series.filter(
      (seriesItem) =>
        normalizeSeriesType(seriesItem.type ?? chart.type) === 'bar' &&
        !seriesItem.stackGroup,
    ).length;
    const barWidth = isBarSeries
      ? resolveBarWidthFromGap(
          item.gapWidth ?? chart.gapWidth,
          barSeriesCount,
          item.overlap ?? chart.overlap,
        )
      : undefined;
    const labelConfig = buildCartesianDataLabelConfig(
      item.dataLabels ?? chart.dataLabels,
      chart.showDataLabels,
      categories,
      isHorizontalBar,
    );

    return {
      name: item.name,
      type: seriesType,
      stack: item.stackGroup,
      data:
        isScatter || isBubbleSeries
          ? item.values.map((value, valueIndex) => [
              item.xValues?.[valueIndex] ??
                categories[valueIndex] ??
                valueIndex + 1,
              value,
            ])
          : item.values,
      areaStyle: item.type === 'area' ? { opacity: 0.18 } : undefined,
      smooth: item.smooth ?? (item.type === 'line' || item.type === 'area'),
      itemStyle: {
        color,
        borderColor: isBarSeries ? '#fff' : color,
        borderWidth: isBarSeries ? 1 : 0,
      },
      lineStyle: {
        color,
        width: item.lineWidth ?? (isLineSeries || item.type === 'area' ? 2 : 1),
      },
      emphasis: {
        itemStyle: {
          color,
          borderColor: isBarSeries ? '#fff' : color,
          borderWidth: isBarSeries ? 1 : 0,
          shadowBlur: isBarSeries ? 6 : 0,
          shadowColor: 'rgba(15, 23, 42, 0.18)',
        },
        lineStyle: {
          color,
          width: item.lineWidth
            ? item.lineWidth + 1
            : isLineSeries || item.type === 'area'
            ? 3
            : 1,
        },
      },
      showSymbol: hideSymbol
        ? false
        : markerSymbol
        ? true
        : isLineSeries ||
          item.type === 'area' ||
          isBubbleSeries ||
          seriesType === 'scatter',
      symbol: markerSymbol,
      symbolSize:
        isBubbleSeries && item.bubbleSizes?.length
          ? (
              _value: unknown,
              parameters: {
                /** 当前结构 在所属集合中的位置索引。 */ dataIndex: number;
              },
            ) =>
              Math.max(
                6,
                Math.min(
                  42,
                  Math.sqrt(
                    Math.abs(item.bubbleSizes?.[parameters.dataIndex] ?? 0),
                  ) * 6,
                ),
              )
          : item.marker?.size ?? (isBubbleSeries ? 14 : 8),
      label: labelConfig,
      barWidth,
      barGap:
        isBarSeries && item.overlap !== undefined
          ? `${-item.overlap}%`
          : undefined,
      barCategoryGap:
        isBarSeries && item.gapWidth !== undefined
          ? `${item.gapWidth}%`
          : undefined,
      barMaxWidth: isBarSeries && !barWidth ? 32 : undefined,
    };
  }) as EChartsOption['series'];

  return {
    animation: false,
    backgroundColor: '#fff',
    color: palette,
    textStyle: OFFICE_TEXT_STYLE,
    title: chart.title
      ? {
          text: chart.title,
          left: 'center',
          top: 8,
          textStyle: {
            fontSize: 14,
            fontWeight: 600,
            color: '#111827',
            fontFamily: OFFICE_FONT_FAMILY,
          },
        }
      : undefined,
    tooltip: {
      trigger: isScatter ? 'item' : 'axis',
      confine: true,
      appendToBody: true,
      axisPointer: isScatter
        ? undefined
        : {
            type: isHorizontalBar ? 'shadow' : 'line',
          },
      backgroundColor: 'rgba(15, 23, 42, 0.96)',
      borderColor: 'rgba(15, 23, 42, 0.96)',
      textStyle: {
        color: '#fff',
        fontFamily: OFFICE_FONT_FAMILY,
      },
    },
    legend: buildLegend(chart),
    grid: buildChartGrid(chart),
    xAxis: isHorizontalBar
      ? {
          type: 'value',
          axisLine: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          axisTick: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          splitLine: {
            lineStyle: {
              color: '#eef2f7',
            },
          },
          axisLabel: {
            hideOverlap: true,
            color: '#475569',
          },
        }
      : {
          type: 'category',
          data: categories,
          axisLine: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          axisTick: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          axisLabel: {
            hideOverlap: true,
            color: '#475569',
          },
        },
    yAxis: isHorizontalBar
      ? {
          type: 'category',
          data: categories,
          axisLine: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          axisTick: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          axisLabel: {
            hideOverlap: true,
            color: '#475569',
          },
        }
      : {
          type: 'value',
          axisLine: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          axisTick: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          splitLine: {
            lineStyle: {
              color: '#eef2f7',
            },
          },
          axisLabel: {
            color: '#475569',
          },
        },
    series,
  };
}

/** 根据输入构建 `buildPieItemStyle` 返回的标准化结果。 */
function buildPieItemStyle(
  series: OfficeChartSeries | undefined,
  index: number,
  palette: string[],
) {
  const pointStyle = series?.pointStyles?.[index];
  const fallbackColor =
    series?.pointColors?.[index] ??
    (series
      ? resolveSeriesColor(series, index)
      : palette[index % palette.length]);
  const color = pointStyle?.color ?? fallbackColor;
  const itemStyle: Record<string, unknown> = {
    color,
  };
  if (pointStyle?.borderColor !== undefined) {
    itemStyle.borderColor = pointStyle.borderColor;
  }
  if (pointStyle?.borderWidth !== undefined) {
    itemStyle.borderWidth = pointStyle.borderWidth;
  }
  return itemStyle;
}
