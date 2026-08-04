import type { EChartsOption } from 'echarts/types/dist/option';
import {
  OFFICE_FONT_FAMILY,
  OFFICE_TEXT_STYLE,
  escapeTooltipHtml,
  resolveCategories,
  sanitizeMapRegionName,
} from './officeChartOptionShared';
import type { OfficeChartModel } from './officeChartTypes';

/** 构建地图图表配置。 */
export function buildMapChartOption(chart: OfficeChartModel): EChartsOption {
  const categories = resolveCategories(chart);
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
          /** 图表回调参数携带的原始或标准化数据。 */
          data?: {
            /** 地图数据项所属的层级名称。 */
            tierName?: string;
          };
          /** 当前图表数据项的显示名称。 */
          name?: string;
          /** 图表回调参数携带的数据值。 */
          value?: unknown;
        };
        const value = typeof item.value === 'number' ? item.value : '';
        const itemName = escapeTooltipHtml(item.name ?? '');
        const seriesName = escapeTooltipHtml(
          sourceSeries?.name ?? chart.mapSeriesName ?? '',
        );
        const tier = item.data?.tierName
          ? `<br/>${escapeTooltipHtml(item.data.tierName)}`
          : '';
        return `${itemName}<br/>${seriesName}: ${value}${tier}`;
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
              /** 图表回调参数携带的原始或标准化数据。 */
              data?: {
                /** 地图数据项的显示标签。 */
                labelName?: string;
              };
              /** 当前图表数据项的显示名称。 */
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
