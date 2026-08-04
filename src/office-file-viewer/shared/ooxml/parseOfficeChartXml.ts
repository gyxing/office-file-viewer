import type {
  OfficeChartColor,
  OfficeChartColorStop,
  OfficeChartModel,
  OfficeChartSeries,
  OfficeChartType,
  OfficeDataLabels,
} from '../chart/officeChartTypes';
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
/** OOXML 图表节点名称到标准图表类型的映射。 */
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

/** 尝试恢复被误按单字节编码解释的 UTF-8 文本。 */
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

function firstText(node: Element | null | undefined) {
  const value =
    textContent(descendantByLocalName(node, 't')) ||
    textContent(descendantByLocalName(node, 'v'));
  return decodeMojibake(value.trim());
}

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

function readFillColor(node: Element | null | undefined, theme: OfficeTheme) {
  const fillNode =
    childByLocalName(node, 'solidFill') ?? childByLocalName(node, 'gradFill');
  const color = readFillValue(fillNode, theme);
  return typeof color === 'string' ? color : undefined;
}

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

function localName(node: Element | null | undefined) {
  return (
    node?.localName.split(':').pop() ??
    node?.localName ??
    ''
  ).toLowerCase();
}

function clamp01(value: number) {
  return Math.max(0, Math.min(1, value));
}

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

function readLineWidth(node: Element | null | undefined) {
  const width = Number(attr(node, 'w'));
  if (!Number.isFinite(width) || width <= 0) return undefined;
  return width / 9525;
}

function readPositiveNumber(node: Element | null | undefined) {
  const value = Number(attr(node, 'val'));
  return Number.isFinite(value) && value > 0 ? value : undefined;
}

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

function isPieLikeChart(type: OfficeChartType, ofPieType?: 'bar' | 'pie') {
  return type === 'pie' || type === 'doughnut' || Boolean(ofPieType);
}

function readShowDataLabels(chartNode: Element | null) {
  return descendantsByLocalName(chartNode, 'dLbls').some((labelsNode) => {
    const showVal = childByLocalName(labelsNode, 'showVal');
    return attr(showVal, 'val') === '1' || attr(showVal, 'val') === 'true';
  });
}

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

function readLegendPosition(chartNode: Element | null) {
  const value = attr(childByLocalName(chartNode, 'legendPos'), 'val');
  if (value === 'b') return 'bottom';
  if (value === 'l') return 'left';
  if (value === 'r') return 'right';
  return 'top';
}

function readFontFamily(node: Element | null | undefined) {
  const latin = attr(childByLocalName(node, 'latin'), 'typeface');
  const eastAsia = attr(childByLocalName(node, 'ea'), 'typeface');
  const complex = attr(childByLocalName(node, 'cs'), 'typeface');
  const value = [eastAsia, latin, complex]
    .filter((item) => item && !item.startsWith('+'))
    .join(', ');
  return value || undefined;
}

function readLegendVisible(chartNode: Element | null) {
  if (!chartNode) return false;
  const deleted = attr(childByLocalName(chartNode, 'delete'), 'val');
  return deleted !== '1' && deleted !== 'true';
}

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

function readSeriesMarker(seriesNode: Element) {
  const markerNode = childByLocalName(seriesNode, 'marker');
  const symbol = attr(childByLocalName(markerNode, 'symbol'), 'val');
  const size = Number(attr(childByLocalName(markerNode, 'size'), 'val'));
  return {
    symbol: symbol ? symbol.toLowerCase() : undefined,
    size: Number.isFinite(size) && size > 0 ? size : undefined,
  };
}

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

function niceRadarMax(value: number) {
  if (!Number.isFinite(value) || value <= 0) return 1;
  const magnitude = 10 ** Math.floor(Math.log10(value));
  const normalized = value / magnitude;
  const step = normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  return step * magnitude;
}

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

/** 将 OOXML 图表部件解析为标准图表模型。 */
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
