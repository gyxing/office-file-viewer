// DocxChartBlock 渲染 DOCX 文档中的独立图表块。
import React, { memo } from 'react';
import type { DocxChartBlock as DocxChartBlockModel } from '../../services/docx/types';
import { DocxChartView } from './DocxChartView';

/** DOCX图表内容块组件属性。 */
type DocxChartBlockProps = {
  /** 当前负责处理或渲染的内容块。 */
  block: DocxChartBlockModel;
  /** 当前预览缩放比例。 */
  zoom: number;
};

/** 渲染DOCX图表内容块。 */
function DocxChartBlockComponent({ block, zoom }: DocxChartBlockProps) {
  return (
    <div className="office-file-docx-chart-block">
      <DocxChartView block={block} zoom={zoom} />
    </div>
  );
}

export const DocxChartBlock = memo(DocxChartBlockComponent);
