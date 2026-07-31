// DocListBlock 渲染 DOC 有序或无序列表块。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { DocListBlock as DocListBlockModel } from '../../services/doc/types';
import { DocInlineContent } from './DocInlineContent';
import { docTextStyleToCss } from './docRenderUtils';

/** DOC列表内容块组件属性。 */
type DocListBlockProps = {
  /** 当前负责处理或渲染的内容块。 */
  block: DocListBlockModel;
};

/** 渲染DOC列表内容块。 */
function DocListBlockComponent({ block }: DocListBlockProps) {
  const itemStyle = useMemo<CSSProperties>(
    () => ({
      ...docTextStyleToCss(block.style),
    }),
    [block.style],
  );
  const Tag = block.ordered ? 'ol' : 'ul';

  return (
    <Tag className="office-file-doc-list">
      {block.items.map((item) => (
        <li
          key={item.id}
          className="office-file-doc-list__item"
          style={itemStyle}
        >
          <DocInlineContent inlines={item.inlines} fallback={item.text} />
        </li>
      ))}
    </Tag>
  );
}

export const DocListBlock = memo(DocListBlockComponent);
