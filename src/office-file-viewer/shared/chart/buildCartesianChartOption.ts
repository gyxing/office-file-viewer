import type { EChartsOption } from 'echarts/types/dist/option';
import {
  OFFICE_FONT_FAMILY,
  OFFICE_TEXT_STYLE,
  buildCartesianDataLabelConfig,
  buildChartGrid,
  buildLegend,
  normalizeSeriesType,
  resolveBarWidthFromGap,
  resolveCategories,
  resolveSeriesColor,
} from './officeChartOptionShared';
import type { OfficeChartModel } from './officeChartTypes';

/** 构建柱形、条形、折线、面积、散点、气泡和组合图配置。 */
export function buildCartesianChartOption(
  chart: OfficeChartModel,
): EChartsOption {
  const categories = resolveCategories(chart);
  const normalizedSeriesTypes = chart.series.map((item) =>
    normalizeSeriesType(item.type ?? chart.type),
  );
  const isHorizontalBar = chart.type === 'bar';
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
