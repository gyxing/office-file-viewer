import type { EChartsOption } from 'echarts/types/dist/option';
import {
  OFFICE_FONT_FAMILY,
  OFFICE_TEXT_STYLE,
  buildLegend,
  buildOfficeTitle,
  buildPieDataLabelConfig,
  buildPieItemStyle,
  escapeTooltipHtml,
  normalizeSeriesType,
  resolveCategories,
  resolveOfficePieStartAngle,
  resolveSeriesColor,
  scaleRoseRadius,
} from './officeChartOptionShared';
import type { OfficeChartModel } from './officeChartTypes';

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
        formatter: `${otherName}<br/>${escapeTooltipHtml(
          sourceSeries?.name ?? '',
        )}: ${secondaryTotal}`,
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
        /** 当前图表组件的子类型。 */
        componentSubType?: string;
        /** 图表回调参数携带的原始或标准化数据。 */
        data?: {
          /** 当前图表数据项的显示名称。 */
          name?: string;
        };
        /** 当前图表数据项的显示名称。 */
        name?: string;
        /** 当前数据项所属的系列名称。 */
        seriesName?: string;
        /** 图表回调参数携带的数据值。 */
        value?: unknown;
      };
      const value =
        typeof item.value === 'number'
          ? item.value
          : Array.isArray(item.value)
          ? item.value[0]
          : '';
      const rawName =
        item.componentSubType === 'bar'
          ? item.seriesName
          : item.data?.name ?? item.name ?? '';
      const name = escapeTooltipHtml(rawName ?? '');
      const seriesName = escapeTooltipHtml(sourceSeries?.name ?? '');
      const valueText =
        typeof value === 'number'
          ? String(value)
          : escapeTooltipHtml(String(value ?? ''));
      return `${name}<br/>${seriesName}: ${valueText}`;
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

/** 构建饼图、环形图和复合饼图配置。 */
export function buildPieChartOption(
  chart: OfficeChartModel,
): EChartsOption | undefined {
  const categories = resolveCategories(chart);
  const normalizedSeriesTypes = chart.series.map((item) =>
    normalizeSeriesType(item.type ?? chart.type),
  );
  const usesMixedSeriesTypes = new Set(normalizedSeriesTypes).size > 1;
  const palette = chart.series.map((series, index) =>
    resolveSeriesColor(series, index),
  );
  const isPie = chart.type === 'pie' || chart.type === 'doughnut';
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

  return undefined;
}
