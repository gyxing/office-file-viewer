// DocxBlockRenderer 根据 DOCX 块类型分发到段落、表格或图表渲染组件。
import React, { memo } from 'react';
import type { DocxBlock } from '../../services/docx/types';
import { DocxChartBlock } from './DocxChartBlock';
import { DocxParagraph } from './DocxParagraph';
import { DocxTableBlock } from './DocxTableBlock';

/** 定义 DocxBlockRenderer 组件可接收的属性。 */
type DocxBlockRendererProps = {
  /** DocxBlockRendererProps 当前负责渲染的文档块模型。 */
  block: DocxBlock;
  /** DocxBlockRendererProps 的 availableWidth 尺寸或坐标，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  availableWidth?: number;
};

/** 渲染 DocxBlockRendererComponent 组件。 */
function DocxBlockRendererComponent({
  block,
  availableWidth,
}: DocxBlockRendererProps) {
  if (block.type === 'table')
    return <DocxTableBlock block={block} availableWidth={availableWidth} />;
  if (block.type === 'chart')
    return <DocxChartBlock block={block} zoom={100} />;
  return <DocxParagraph block={block} />;
}

export const DocxBlockRenderer = memo(DocxBlockRendererComponent);
