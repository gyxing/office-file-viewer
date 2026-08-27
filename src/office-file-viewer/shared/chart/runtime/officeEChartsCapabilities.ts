import type { OfficeChartModel, OfficeChartType } from '../officeChartTypes';

/** Office 图表按需加载时采用的稳定能力组。 */
export type OfficeEChartsCapability =
  | 'cartesian'
  | 'pie'
  | 'radar'
  | 'map'
  | 'labels';

/** 判断当前图表是否真的需要加载数据标签布局能力。 */
function hasVisibleDataLabels(chart: OfficeChartModel) {
  if (chart.showDataLabels) return true;
  return [chart.dataLabels, ...chart.series.map((series) => series.dataLabels)]
    .filter(Boolean)
    .some(
      (labels) =>
        !labels?.delete &&
        Boolean(
          labels?.showVal ||
            labels?.showCatName ||
            labels?.showSerName ||
            labels?.showPercent,
        ),
    );
}

/** 把标准图表类型映射到对应的 ECharts 能力组。 */
function resolveChartTypeCapability(
  type: OfficeChartType,
): OfficeEChartsCapability {
  if (type === 'pie' || type === 'doughnut') return 'pie';
  if (type === 'radar') return 'radar';
  if (type === 'map') return 'map';
  return 'cartesian';
}

/** 解析当前图表真正需要注册的稳定 ECharts 能力组。 */
export function resolveOfficeEChartsCapabilities(
  chart: OfficeChartModel,
): OfficeEChartsCapability[] {
  if (!chart.series.length) return [];

  const capabilities = new Set<OfficeEChartsCapability>([
    resolveChartTypeCapability(chart.type),
  ]);
  chart.series.forEach((series) => {
    if (series.type) {
      capabilities.add(resolveChartTypeCapability(series.type));
    }
  });

  // 复合饼图的第二绘图区为条形时，还需要笛卡尔坐标与柱状图能力。
  if (
    (chart.type === 'pie' || chart.type === 'doughnut') &&
    chart.ofPieType === 'bar'
  ) {
    capabilities.add('cartesian');
  }

  if (hasVisibleDataLabels(chart)) capabilities.add('labels');

  return Array.from(capabilities);
}
