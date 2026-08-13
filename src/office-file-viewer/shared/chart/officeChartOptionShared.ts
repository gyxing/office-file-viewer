import type {
  OfficeChartModel,
  OfficeChartSeries,
  OfficeChartType,
  OfficeDataLabels,
} from './officeChartTypes';

/** Office 图表没有显式配色时使用的默认颜色序列。 */
export const DEFAULT_COLORS = [
  '#5470c6',
  '#91cc75',
  '#fac858',
  '#ee6666',
  '#73c0de',
  '#3ba272',
  '#fc8452',
];
/** Office 图表缺少字体信息时使用的默认字体回退栈。 */
export const OFFICE_FONT_FAMILY =
  '"Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", Arial, sans-serif';
/** Office 图表文字使用的默认颜色和字体样式。 */
export const OFFICE_TEXT_STYLE = {
  color: '#334155',
  fontFamily: OFFICE_FONT_FAMILY,
};
/** ECharts HTML Tooltip 中需要转义的字符及其实体。 */
const TOOLTIP_HTML_ENTITIES: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;',
};

/** 转义来自 Office 文档的 Tooltip 文本，避免被当作 HTML 解析。 */
export function escapeTooltipHtml(value: string) {
  return value.replace(
    /[&<>"']/g,
    (character) => TOOLTIP_HTML_ENTITIES[character] ?? character,
  );
}

function niceRadarMax(value: number) {
  if (!Number.isFinite(value) || value <= 0) return 1;
  const magnitude = 10 ** Math.floor(Math.log10(value));
  const normalized = value / magnitude;
  const step = normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  return step * magnitude;
}

/** 按所有系列的维度最大值生成雷达轴范围。 */
export function buildRadarIndicators(
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

/** 优先使用系列显式颜色，否则按索引选用默认配色。 */
export function resolveSeriesColor(series: OfficeChartSeries, index: number) {
  return series.color ?? DEFAULT_COLORS[index % DEFAULT_COLORS.length];
}

/** 返回显式分类，缺失时按最长系列生成一基序号。 */
export function resolveCategories(chart: OfficeChartModel) {
  if (chart.categories.length) {
    return chart.categories;
  }

  const maxLength = Math.max(
    ...chart.series.map((series) => series.values.length),
    0,
  );
  return Array.from({ length: maxLength }, (_, index) => String(index + 1));
}

/** 将 Office 图表类型映射为 ECharts 系列类型。 */
export function normalizeSeriesType(type: OfficeChartType) {
  if (type === 'column' || type === 'bar') return 'bar';
  if (type === 'area') return 'line';
  if (type === 'scatter' || type === 'bubble') return 'scatter';
  if (type === 'radar') return 'radar';
  if (type === 'pie' || type === 'doughnut') return 'pie';
  return 'line';
}

/** 去除地图区域名称中的行政区后缀。 */
export function sanitizeMapRegionName(name: string) {
  return name
    .replace(
      /特别行政区$|壮族自治区$|回族自治区$|维吾尔自治区$|自治区$|省$|市$/g,
      '',
    )
    .trim();
}

/** 将玫瑰图半径约束到预览容器的可见范围。 */
export function scaleRoseRadius(
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

/** 生成各图表变体共用的图例配置。 */
export function buildLegend(
  chart: OfficeChartModel,
  itemCount = chart.series.length,
) {
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
        top: chart.title ? 45 : 8,
      };
  }
}

/** 根据标题、图例和坐标轴可见性计算笛卡尔绘图区边距。 */
export function buildChartGrid(chart: OfficeChartModel) {
  const isBottomLegend = chart.legendPosition === 'bottom';
  const isSideLegend =
    chart.legendPosition === 'left' || chart.legendPosition === 'right';
  const hasTopLegend =
    chart.showLegend !== false &&
    (!chart.legendPosition || chart.legendPosition === 'top');
  const hidesVerticalValueAxis =
    chart.type !== 'bar' && chart.showValueAxis === false;
  return {
    left: isSideLegend ? 70 : hidesVerticalValueAxis ? 15 : 40,
    right: isSideLegend ? 70 : hidesVerticalValueAxis ? 15 : 24,
    // 标题和顶部图例各占一行，绘图区不能与图例重叠。
    top: chart.title ? (hasTopLegend ? 82 : 56) : hasTopLegend ? 40 : 24,
    // containLabel 会额外计入分类轴标签高度，普通底边距无需再次重复预留。
    bottom: isBottomLegend ? 56 : 12,
    containLabel: true,
  };
}

/** 生成各图表变体共用的标题配置。 */
export function buildOfficeTitle(chart: OfficeChartModel) {
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

/** 将 Office 饼图起始角转换为 ECharts 角度。 */
export function resolveOfficePieStartAngle(chart: OfficeChartModel) {
  const officeAngle = chart.startAngle ?? 0;
  return (((90 - officeAngle) % 360) + 360) % 360;
}

/** 返回雷达图显式起始角或 Office 默认角度。 */
export function resolveOfficeRadarStartAngle(chart: OfficeChartModel) {
  return chart.radarStartAngle ?? 90;
}

/** 根据图例位置确定雷达图显示半径。 */
export function resolveOfficeRadarRadius(chart: OfficeChartModel) {
  return (
    chart.radarRadius ?? (chart.legendPosition === 'bottom' ? '62%' : '68%')
  );
}

/** 按 Office 的顺时针轴顺序重排雷达维度。 */
export function reorderRadarAxes<T>(items: T[]) {
  if (items.length <= 2) return items.slice();
  return [items[0], ...items.slice(1).reverse()];
}

/** 根据标题和图例位置确定雷达图中心。 */
export function resolveOfficeRadarCenter(
  chart: OfficeChartModel,
): [string, string] {
  if (chart.legendPosition === 'bottom')
    return ['50%', chart.title ? '52%' : '48%'];
  if (chart.legendPosition === 'top')
    return ['50%', chart.title ? '58%' : '56%'];
  return ['50%', chart.title ? '56%' : '52%'];
}

/** 将 Office 以柱宽为基准的分类间距换算为 ECharts 分类带宽百分比。 */
export function resolveBarCategoryGap(
  gapWidth: number | undefined,
  seriesCount: number,
  overlap?: number,
) {
  if (!Number.isFinite(gapWidth) || gapWidth === undefined) return undefined;
  const visibleSeriesCount = Math.max(1, seriesCount);
  const overlapRatio = Math.max(-1, Math.min(1, (overlap ?? 0) / 100));
  const groupUnits = Math.max(
    1,
    visibleSeriesCount - overlapRatio * (visibleSeriesCount - 1),
  );
  const gapUnits = Math.max(0, gapWidth / 100);
  const categoryGap = (gapUnits / (groupUnits + gapUnits)) * 100;
  return `${Math.min(95, Math.round(categoryGap * 1000) / 1000)}%`;
}

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

/** 按标签显示规则组合系列、分类、数值和百分比。 */
export function buildDataLabelFormatter(
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
      /** 当前图表数据项的显示名称。 */
      name?: string;
      /** 图表回调参数携带的数据值。 */
      value?: unknown;
      /** 当前数据项在所属系列中的零基索引。 */
      dataIndex?: number;
      /** 面向界面展示的解析完成百分比。 */
      percent?: number;
      /** 当前数据项所属的系列名称。 */
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

/** 合并图表级与系列级规则，判断是否显示数据标签。 */
export function shouldShowDataLabels(
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

/** 生成笛卡尔图表的数据标签配置。 */
export function buildCartesianDataLabelConfig(
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

/** 生成饼图与环形图的数据标签配置。 */
export function buildPieDataLabelConfig(
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
      /** 当前图表数据项的显示名称。 */
      name?: string;
      /** 图表回调参数携带的原始或标准化数据。 */
      data?: {
        /** 当前图表数据项的显示名称。 */
        name?: string;
      };
      /** 图表回调参数携带的数据值。 */
      value?: number;
      /** 面向界面展示的解析完成百分比。 */
      percent?: number;
      /** 当前数据项所属的系列名称。 */
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
    color: position === 'inside' ? '#ffffff' : '#334155',
    fontSize: 12,
    fontFamily: OFFICE_FONT_FAMILY,
  };
}

/** 合并数据点、系列和默认调色板得到扇区样式。 */
export function buildPieItemStyle(
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
