// DocBlockRenderer 根据 DOC 块类型分发到段落、列表或表格渲染组件。
import React, { memo } from 'react';
import type { DocBlock } from '../../services/doc/types';
import { DocListBlock } from './DocListBlock';
import { DocParagraphBlock } from './DocParagraphBlock';
import { DocTableBlock } from './DocTableBlock';

/** 定义 DocBlockRenderer 组件可接收的属性。 */
type DocBlockRendererProps = {
  /** DocBlockRendererProps 当前负责渲染的文档块模型。 */
  block: DocBlock;
};

/** 渲染 DocBlockRendererComponent 组件。 */
function DocBlockRendererComponent({ block }: DocBlockRendererProps) {
  if (block.type === 'table') return <DocTableBlock block={block} />;
  if (block.type === 'list') return <DocListBlock block={block} />;
  return <DocParagraphBlock block={block} />;
}

export const DocBlockRenderer = memo(DocBlockRendererComponent);
