// DocxInlineChart 渲染 DOCX 行内图表。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { DocxInline } from '../../services/docx/types';
import { DocxChartView } from './DocxChartView';

/** DOCX行内内容图表组件属性。 */
type DocxInlineChartProps = {
  /** 当前负责渲染的行内内容模型。 */
  inline: Extract<
    DocxInline,
    {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'chart';
    }
  >;
};

/** 渲染DOCX行内内容图表。 */
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
      <DocxChartView block={chart} zoom={100} />
    </span>
  );
}

export const DocxInlineChart = memo(DocxInlineChartComponent);
