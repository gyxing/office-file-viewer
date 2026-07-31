// DocxBlockRenderer 根据 DOCX 块类型分发到段落、表格或图表渲染组件。
import React, { memo } from 'react';
import type { DocxBlock } from '../../services/docx/types';
import { DocxChartBlock } from './DocxChartBlock';
import { DocxParagraph } from './DocxParagraph';
import { DocxTableBlock } from './DocxTableBlock';

/** DOCX内容块渲染器组件属性。 */
type DocxBlockRendererProps = {
  /** 当前负责处理或渲染的内容块。 */
  block: DocxBlock;
  /** 当前可用宽度，单位为标准化渲染像素。 */
  availableWidth?: number;
  /** 当前页面或容器允许块内容占用的最大物理宽度。 */
  maximumWidth?: number;
};

/** 渲染DOCX内容块渲染器。 */
function DocxBlockRendererComponent({
  block,
  availableWidth,
  maximumWidth,
}: DocxBlockRendererProps) {
  if (block.type === 'table')
    return (
      <DocxTableBlock
        block={block}
        availableWidth={availableWidth}
        maximumWidth={maximumWidth}
      />
    );
  if (block.type === 'chart')
    return <DocxChartBlock block={block} zoom={100} />;
  return <DocxParagraph block={block} />;
}

export const DocxBlockRenderer = memo(DocxBlockRendererComponent);
