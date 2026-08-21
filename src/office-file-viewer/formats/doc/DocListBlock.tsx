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
  const { itemStyle, listStyle, markerWidth } = useMemo(() => {
    const sourceStyle = docTextStyleToCss(block.style, resolveFontFamily);
    const indentLeft = Math.max(20, block.style?.indentLeft ?? 0);
    const hangingIndent = Math.max(
      20,
      block.style?.firstLineIndent && block.style.firstLineIndent < 0
        ? -block.style.firstLineIndent
        : 0,
    );
    const markerStart = Math.max(0, indentLeft - hangingIndent);
    return {
      markerWidth: hangingIndent,
      listStyle: {
        ...sourceStyle,
        marginLeft: markerStart,
        marginBottom:
          !block.continuesOnNext && block.style?.spacingAfter === undefined
            ? 16
            : 0,
        paddingLeft: 0,
        textIndent: 0,
        listStyle: 'none',
      } satisfies CSSProperties,
      itemStyle: {
        ...sourceStyle,
        gridTemplateColumns: `${hangingIndent}px minmax(0, 1fr)`,
        marginTop: 0,
        marginRight: 0,
        marginBottom: block.style?.spacingAfter ?? 8,
        marginLeft: 0,
        padding: 0,
        textIndent: 0,
      } satisfies CSSProperties,
    };
  }, [block.style, resolveFontFamily]);
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
      style={listStyle}
    >
      {block.items.map((item, itemIndex) => (
        <li
          key={item.id}
          className="office-file-doc-list__item"
          style={itemStyle}
        >
          <span
            className="office-file-doc-list__marker"
            aria-hidden="true"
            style={{ width: markerWidth }}
          >
            {item.marker ?? (block.ordered ? `${itemIndex + 1}.` : '•')}
          </span>
          <span className="office-file-doc-list__content">
            <DocInlineContent
              inlines={item.inlines}
              fallback={item.text}
              sourceId={item.id}
              searchBlockId={searchBlockId}
            />
          </span>
        </li>
      ))}
    </Tag>
  );
}

export const DocListBlock = memo(DocListBlockComponent);
