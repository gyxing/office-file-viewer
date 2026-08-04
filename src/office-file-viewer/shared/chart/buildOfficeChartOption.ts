import type { EChartsOption } from 'echarts/types/dist/option';
import { buildCartesianChartOption } from './buildCartesianChartOption';
import { buildMapChartOption } from './buildMapChartOption';
import { buildPieChartOption } from './buildPieChartOption';
import { buildRadarChartOption } from './buildRadarChartOption';
import type { OfficeChartModel } from './officeChartTypes';

/** 将标准图表模型穷尽分派到对应的 ECharts 配置构建器。 */
export function buildOfficeChartOption(chart: OfficeChartModel): EChartsOption {
  if (!chart.series.length) {
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

  const chartType = chart.type;
  switch (chartType) {
    case 'map':
      return buildMapChartOption(chart);
    case 'radar':
      return buildRadarChartOption(chart) ?? buildCartesianChartOption(chart);
    case 'pie':
    case 'doughnut':
      return buildPieChartOption(chart) ?? buildCartesianChartOption(chart);
    case 'line':
    case 'bar':
    case 'column':
    case 'area':
    case 'scatter':
    case 'bubble':
    case 'unknown':
      return buildCartesianChartOption(chart);
  }

  const unreachableType: never = chartType;
  throw new Error(`不支持的标准图表类型：${unreachableType}`);
}
