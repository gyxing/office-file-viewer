// DocInlineContent 渲染 DOC 段落或表格单元格内的文本和图片片段。
import React, { memo } from 'react';
import type { DocTextInline } from '../../services/doc/types';
import { inlineStyleToCss } from './docRenderUtils';

/** DOC 行内内容组件属性。 */
type DocInlineContentProps = {
  /** 按源文档顺序排列的行内内容。 */
  inlines?: DocTextInline[];
  /** 没有可渲染行内节点时显示的回退文本。 */
  fallback: string;
  /** 是否保留块级模型自身的字体与段落样式。 */
  preserveBlockTypography?: boolean;
  /** 是否使用 Word 表格的数字分隔符换行规则。 */
  wordTableLineBreaks?: boolean;
};

/** 恢复 Word 表格中斜杠与后续数字不拆行的排版语义。 */
function withWordTableLineBreaks(text: string) {
  return text.replace(/\/(?=\d)/g, '\u200b/\u2060');
}

/** 渲染 DOC 段落中的行内文字和图片。 */
function DocInlineContentComponent({
  inlines,
  fallback,
  preserveBlockTypography,
  wordTableLineBreaks,
}: DocInlineContentProps) {
  if (!inlines?.length) {
    return (
      <>{wordTableLineBreaks ? withWordTableLineBreaks(fallback) : fallback}</>
    );
  }

  return (
    <>
      {inlines.map((inline, index) =>
        inline.type === 'image' ? (
          <span
            key={`${inline.image.id}-${index}`}
            className="office-file-doc-inline-image"
          >
            <img
              className="office-file-doc-inline-image__img"
              src={inline.image.src}
              alt={inline.image.caption ?? inline.image.id}
              loading="lazy"
              decoding="async"
              style={{
                width:
                  inline.image.width && inline.image.width <= 520
                    ? inline.image.width
                    : undefined,
                height:
                  inline.image.height &&
                  inline.image.width &&
                  inline.image.width <= 520
                    ? inline.image.height
                    : undefined,
              }}
            />
          </span>
        ) : (
          <span
            key={`${inline.text}-${index}`}
            style={inlineStyleToCss(inline.style, { preserveBlockTypography })}
          >
            {wordTableLineBreaks
              ? withWordTableLineBreaks(inline.text)
              : inline.text}
          </span>
        ),
      )}
    </>
  );
}

export const DocInlineContent = memo(DocInlineContentComponent);
