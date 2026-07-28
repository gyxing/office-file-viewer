// XlsxFloatingCharts 渲染锚定在工作表画布上的浮动图表。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { XlsxChart, XlsxSheet } from '../../services/xlsx/types';
import { OfficeChartView } from '../../shared/chart/OfficeChartView';
import {
  getXlsxMeasuredAnchorRect,
  XLSX_ROW_HEADER_WIDTH,
  type XlsxMeasuredAnchorRect,
  type XlsxSheetMetrics,
} from './sheetRenderUtils';

/** 定义 XlsxFloatingCharts 组件可接收的属性。 */
type XlsxFloatingChartsProps = {
  /** 当前关联的工作表模型。 */
  sheet: XlsxSheet;
  /** 浏览器最终表格布局对应的工作表指标。 */
  metrics: XlsxSheetMetrics;
};

/** 渲染 XlsxFloatingChart 组件。 */
function XlsxFloatingChart({
  chart,
  rect,
  columnHeaderHeight,
}: {
  /** 当前关联的图表模型。 */
  chart: XlsxChart;
  /** 按浏览器最终表格布局重算的锚点矩形。 */
  rect: XlsxMeasuredAnchorRect;
  /** 浏览器最终计算出的列标题行高度。 */
  columnHeaderHeight: number;
}) {
  const chartStyle = useMemo<CSSProperties>(
    () => ({
      left: XLSX_ROW_HEADER_WIDTH + rect.x,
      top: columnHeaderHeight + rect.y,
      width: rect.width,
      height: rect.height,
    }),
    [columnHeaderHeight, rect.height, rect.width, rect.x, rect.y],
  );

  return (
    <div
      className="office-file-xlsx-sheet-grid__floating-chart"
      style={chartStyle}
    >
      <OfficeChartView
        chart={chart.chart}
        width={rect.width}
        height={rect.height}
        zoom={100}
      />
    </div>
  );
}

const MemoXlsxFloatingChart = memo(XlsxFloatingChart);

/** 渲染 XlsxFloatingChartsComponent 组件。 */
function XlsxFloatingChartsComponent({
  sheet,
  metrics,
}: XlsxFloatingChartsProps) {
  const positionedCharts = useMemo(
    () =>
      sheet.charts.map((chart) => ({
        chart,
        rect: getXlsxMeasuredAnchorRect(sheet, metrics, chart),
      })),
    [metrics, sheet],
  );

  return (
    <>
      {positionedCharts.map(({ chart, rect }) => (
        <MemoXlsxFloatingChart
          key={chart.id}
          chart={chart}
          rect={rect}
          columnHeaderHeight={metrics.columnHeaderHeight}
        />
      ))}
    </>
  );
}

export const XlsxFloatingCharts = memo(XlsxFloatingChartsComponent);
