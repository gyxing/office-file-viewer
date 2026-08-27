import type { EChartsOption } from 'echarts/types/dist/option';
import {
  OFFICE_FONT_FAMILY,
  OFFICE_TEXT_STYLE,
  buildCartesianDataLabelConfig,
  buildChartGrid,
  buildLegend,
  normalizeSeriesType,
  resolveBarCategoryGap,
  resolveCategories,
  resolveSeriesColor,
} from './officeChartOptionShared';
import type { OfficeChartModel } from './officeChartTypes';

/** 判断图表是否需要为轴外数据标签预留一个主刻度。 */
function hasVisibleDataLabels(chart: OfficeChartModel) {
  if (chart.showDataLabels) return true;
  return chart.series.some((series) => {
    const labels = series.dataLabels;
    return (
      !labels?.delete &&
      Boolean(
        labels?.showVal ||
          labels?.showCatName ||
          labels?.showSerName ||
          labels?.showPercent,
      )
    );
  });
}

/** 按 Office 常用的 1/2/5 刻度序列估算自动数值轴上界。 */
function resolveValueAxisMaximum(chart: OfficeChartModel) {
  if (chart.valueAxisMaximum !== undefined) return chart.valueAxisMaximum;
  if (!hasVisibleDataLabels(chart)) return undefined;
  const maximum = Math.max(
    0,
    ...chart.series.flatMap((series) => series.values),
  );
  if (maximum <= 0) return undefined;
  const rawStep = maximum / 5;
  const magnitude = 10 ** Math.floor(Math.log10(rawStep));
  const normalized = rawStep / magnitude;
  const stepMultiplier = normalized <= 1 ? 1 : normalized <= 2 ? 2 : 5;
  const step = stepMultiplier * magnitude;
  const roundedMaximum = Math.ceil(maximum / step) * step;
  // 最大数据点正好压在主刻度上时，Office 会再留一格展示轴外标签。
  return Math.abs(roundedMaximum - maximum) <= step * 1e-9
    ? roundedMaximum + step
    : roundedMaximum;
}

/** 构建柱形、条形、折线、面积、散点、气泡和组合图配置。 */
export function buildCartesianChartOption(
  chart: OfficeChartModel,
): EChartsOption {
  const categories = resolveCategories(chart);
  const normalizedSeriesTypes = chart.series.map((item) =>
    normalizeSeriesType(item.type ?? chart.type),
  );
  const isHorizontalBar = chart.type === 'bar';
  const showCategoryAxis = chart.showCategoryAxis !== false;
  const showValueAxis = chart.showValueAxis !== false;
  const valueAxisMaximum = resolveValueAxisMaximum(chart);
  const isScatter =
    normalizedSeriesTypes.length > 0 &&
    normalizedSeriesTypes.every((type) => type === 'scatter');
  const palette = chart.series.map((series, index) =>
    resolveSeriesColor(series, index),
  );
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
    const barCategoryGap = isBarSeries
      ? resolveBarCategoryGap(
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
        ? labelConfig.show
        : markerSymbol
        ? true
        : isLineSeries ||
          item.type === 'area' ||
          isBubbleSeries ||
          seriesType === 'scatter',
      // ECharts 的折线标签依附数据点符号；源文件隐藏标记但保留标签时使用零尺寸锚点。
      symbol: hideSymbol && labelConfig.show ? 'circle' : markerSymbol,
      symbolSize:
        hideSymbol && labelConfig.show
          ? 0
          : isBubbleSeries && item.bubbleSizes?.length
          ? (
              _value: unknown,
              parameters: {
                /** 当前数据项在所属系列中的零基索引。 */
                dataIndex: number;
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
      barGap:
        isBarSeries && item.overlap !== undefined
          ? `${-item.overlap}%`
          : undefined,
      barCategoryGap,
      barMaxWidth: isBarSeries && item.gapWidth === undefined ? 32 : undefined,
    };
  }) as EChartsOption['series'];

  const categoryAxisOption = {
    type: 'category' as const,
    data: categories,
    axisLine: {
      show: showCategoryAxis,
      lineStyle: { color: '#cbd5e1' },
    },
    axisTick: {
      show: showCategoryAxis,
      lineStyle: { color: '#cbd5e1' },
    },
    axisLabel: {
      show: showCategoryAxis,
      hideOverlap: true,
      color: '#475569',
    },
  };
  const valueAxisOption = {
    type: 'value' as const,
    min: chart.valueAxisMinimum,
    max: valueAxisMaximum,
    interval: chart.valueAxisMajorUnit,
    axisLine: {
      show: showValueAxis,
      lineStyle: { color: '#cbd5e1' },
    },
    axisTick: {
      show: showValueAxis,
      lineStyle: { color: '#cbd5e1' },
    },
    // Office 可隐藏数值轴但保留主网格线，二者不能共用同一个可见性开关。
    splitLine: {
      lineStyle: { color: '#eef2f7' },
    },
    axisLabel: {
      show: showValueAxis,
      hideOverlap: true,
      color: '#475569',
    },
  };

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
    xAxis: isHorizontalBar ? valueAxisOption : categoryAxisOption,
    yAxis: isHorizontalBar ? categoryAxisOption : valueAxisOption,
    series,
  };
}
