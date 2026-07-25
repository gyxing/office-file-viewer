// DocxChartBlock 渲染 DOCX 文档中的独立图表块。
import React, { memo } from 'react';
import type { DocxChartBlock as DocxChartBlockModel } from '../../services/docx/types';
import { OfficeChartView } from '../../shared/chart/OfficeChartView';

/** 定义 DocxChartBlock 组件可接收的属性。 */
type DocxChartBlockProps = {
  /** DocxChartBlockProps 当前负责渲染的文档块模型。 */
  block: DocxChartBlockModel;
  /** 当前预览缩放比例。 */
  zoom: number;
};

/** 渲染 DocxChartBlockComponent 组件。 */
function DocxChartBlockComponent({ block, zoom }: DocxChartBlockProps) {
  return (
    <div className="office-file-docx-chart-block">
      <OfficeChartView
        chart={block.chart}
        width={block.width}
        height={block.height}
        zoom={zoom}
      />
    </div>
  );
}

export const DocxChartBlock = memo(DocxChartBlockComponent);
