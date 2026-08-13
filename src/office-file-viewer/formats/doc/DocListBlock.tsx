// DocListBlock 渲染 DOC 有序或无序列表块。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { DocListBlock as DocListBlockModel } from '../../services/doc/types';
import { useOfficeFontResolver } from '../../shared/fonts/OfficeFontProvider';
import { DocInlineContent } from './DocInlineContent';
import { docTextStyleToCss } from './docRenderUtils';

/** DOC列表内容块组件属性。 */
type DocListBlockProps = {
  /** 当前负责处理或渲染的内容块。 */
  block: DocListBlockModel;
};

/** 渲染DOC列表内容块。 */
function DocListBlockComponent({ block }: DocListBlockProps) {
  const resolveFontFamily = useOfficeFontResolver();
  const searchBlockId = block.sourceBlockId ?? block.id;
  const itemStyle = useMemo<CSSProperties>(
    () => ({
      ...docTextStyleToCss(block.style, resolveFontFamily),
    }),
    [block.style, resolveFontFamily],
  );
  const Tag = block.ordered ? 'ol' : 'ul';

  return (
    <Tag
      className="office-file-doc-list"
      data-office-word-block-id={searchBlockId}
      data-office-doc-estimated-height={block.estimatedHeight}
      data-office-doc-pagination-id={block.sourceBlockId ?? block.id}
      data-office-doc-pagination-fragment={
        block.sourceBlockId ? 'true' : undefined
      }
    >
      {block.items.map((item) => (
        <li
          key={item.id}
          className="office-file-doc-list__item"
          style={itemStyle}
        >
          <DocInlineContent
            inlines={item.inlines}
            fallback={item.text}
            sourceId={item.id}
            searchBlockId={searchBlockId}
          />
        </li>
      ))}
    </Tag>
  );
}

export const DocListBlock = memo(DocListBlockComponent);
