// XlsxFloatingCharts 渲染锚定在工作表画布上的浮动图表。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { XlsxChart } from '../../services/xlsx/types';
import { OfficeChartView } from '../../shared/chart/OfficeChartView';

/** 定义 XlsxFloatingCharts 组件可接收的属性。 */
type XlsxFloatingChartsProps = {
  /** XlsxFloatingChartsProps 包含的 charts 有序集合。 */
  charts: XlsxChart[];
};

/** 渲染 XlsxFloatingChart 组件。 */
function XlsxFloatingChart({
  chart,
}: {
  /** 当前结构 当前关联的图表模型。 */ chart: XlsxChart;
}) {
  const chartStyle = useMemo<CSSProperties>(
    () => ({
      left: 48 + chart.x,
      top: 28 + chart.y,
      width: chart.width,
      height: chart.height,
    }),
    [chart.height, chart.width, chart.x, chart.y],
  );

  return (
    <div
      className="office-file-xlsx-sheet-grid__floating-chart"
      style={chartStyle}
    >
      <OfficeChartView
        chart={chart.chart}
        width={chart.width}
        height={chart.height}
        zoom={100}
      />
    </div>
  );
}

const MemoXlsxFloatingChart = memo(XlsxFloatingChart);

/** 渲染 XlsxFloatingChartsComponent 组件。 */
function XlsxFloatingChartsComponent({ charts }: XlsxFloatingChartsProps) {
  return (
    <>
      {charts.map((chart) => (
        <MemoXlsxFloatingChart key={chart.id} chart={chart} />
      ))}
    </>
  );
}

export const XlsxFloatingCharts = memo(XlsxFloatingChartsComponent);
