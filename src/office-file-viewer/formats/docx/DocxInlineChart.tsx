// DocxInlineChart 渲染 DOCX 行内图表。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { DocxInline } from '../../services/docx/types';
import { OfficeChartView } from '../../shared/chart/OfficeChartView';

/** 定义 DocxInlineChart 组件可接收的属性。 */
type DocxInlineChartProps = {
  /** DocxInlineChartProps 当前负责渲染的行内内容模型。 */
  inline: Extract<
    DocxInline,
    {
      /** 用于区分 DocxInlineChartProps 不同结构分支的类型标识。 */
      type: 'chart';
    }
  >;
};

/** 渲染 DocxInlineChartComponent 组件。 */
function DocxInlineChartComponent({ inline }: DocxInlineChartProps) {
  const chart = inline.chart;
  const chartStyle = useMemo<CSSProperties>(
    () =>
      ({
        '--office-file-docx-inline-chart-width': `${chart.width}px`,
        '--office-file-docx-inline-chart-height': `${chart.height}px`,
      } as CSSProperties),
    [chart.height, chart.width],
  );

  return (
    <span className="office-file-docx-inline-chart" style={chartStyle}>
      <OfficeChartView
        chart={chart.chart}
        width={chart.width}
        height={chart.height}
        zoom={100}
      />
    </span>
  );
}

export const DocxInlineChart = memo(DocxInlineChartComponent);
