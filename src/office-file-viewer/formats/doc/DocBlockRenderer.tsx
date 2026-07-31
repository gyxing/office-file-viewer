// DocBlockRenderer 根据 DOC 块类型分发到段落、列表或表格渲染组件。
import React, { memo } from 'react';
import type { DocBlock } from '../../services/doc/types';
import { DocListBlock } from './DocListBlock';
import { DocParagraphBlock } from './DocParagraphBlock';
import { DocTableBlock } from './DocTableBlock';

/** DOC内容块渲染器组件属性。 */
type DocBlockRendererProps = {
  /** 当前负责处理或渲染的内容块。 */
  block: DocBlock;
};

/** 渲染DOC内容块渲染器。 */
function DocBlockRendererComponent({ block }: DocBlockRendererProps) {
  if (block.type === 'table') return <DocTableBlock block={block} />;
  if (block.type === 'list') return <DocListBlock block={block} />;
  return <DocParagraphBlock block={block} />;
}

export const DocBlockRenderer = memo(DocBlockRendererComponent);
