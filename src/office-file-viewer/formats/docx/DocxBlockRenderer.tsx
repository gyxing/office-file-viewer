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
  /** 查找结果对应的顶层正文块标识。 */
  searchBlockId?: string;
  /** 是否忽略相邻同样式段落之间的段前距。 */
  suppressSpacingBefore?: boolean;
  /** 覆盖浏览器实际应用的段前距。 */
  spacingBefore?: number;
  /** 是否忽略相邻同样式段落之间的段后距。 */
  suppressSpacingAfter?: boolean;
};

/** 渲染DOCX内容块渲染器。 */
function DocxBlockRendererComponent({
  block,
  availableWidth,
  maximumWidth,
  searchBlockId = block.sourceBlockId ?? block.id,
  suppressSpacingBefore = false,
  suppressSpacingAfter = false,
  spacingBefore,
}: DocxBlockRendererProps) {
  if (block.type === 'table')
    return (
      <DocxTableBlock
        block={block}
        availableWidth={availableWidth}
        maximumWidth={maximumWidth}
        searchBlockId={searchBlockId}
      />
    );
  if (block.type === 'chart')
    return <DocxChartBlock block={block} zoom={100} />;
  return (
    <DocxParagraph
      block={block}
      searchBlockId={searchBlockId}
      suppressSpacingBefore={suppressSpacingBefore}
      suppressSpacingAfter={suppressSpacingAfter}
      spacingBefore={spacingBefore}
    />
  );
}

export const DocxBlockRenderer = memo(DocxBlockRendererComponent);
