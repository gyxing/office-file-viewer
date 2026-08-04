import type { EChartsOption } from 'echarts/types/dist/option';
import {
  OFFICE_FONT_FAMILY,
  OFFICE_TEXT_STYLE,
  buildDataLabelFormatter,
  buildLegend,
  buildRadarIndicators,
  reorderRadarAxes,
  resolveCategories,
  resolveOfficeRadarCenter,
  resolveOfficeRadarRadius,
  resolveOfficeRadarStartAngle,
  resolveSeriesColor,
  shouldShowDataLabels,
} from './officeChartOptionShared';
import type { OfficeChartModel } from './officeChartTypes';

/** 构建雷达图配置；缺少有效维度时由总入口降级到笛卡尔配置。 */
export function buildRadarChartOption(
  chart: OfficeChartModel,
): EChartsOption | undefined {
  const categories = resolveCategories(chart);
  const palette = chart.series.map((series, index) =>
    resolveSeriesColor(series, index),
  );
  const radarIndicators =
    chart.radarIndicators ?? buildRadarIndicators(categories, chart.series);
  const radarDisplayIndicators = radarIndicators.length
    ? reorderRadarAxes(radarIndicators)
    : undefined;
  if (!radarDisplayIndicators?.length) return undefined;
  const radarCategories = radarDisplayIndicators.map(
    (indicator) => indicator.name,
  );
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
