// DocParagraphBlock 渲染 DOC 段落块，并应用推断出的标题、正文和文字样式。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type {
  DocParagraphBlock as DocParagraphBlockModel,
  DocTextInline,
} from '../../services/doc/types';
import { DocInlineContent } from './DocInlineContent';
import { docTextStyleToCss } from './docRenderUtils';

/** DOC段落内容块组件属性。 */
type DocParagraphBlockProps = {
  /** 当前负责处理或渲染的内容块。 */
  block: DocParagraphBlockModel;
};

/** 按目录段落最后一个制表符拆出标题和页码，并保留各自的源文字样式。 */
function splitTableOfContentsInlines(inlines?: DocTextInline[]) {
  if (!inlines?.length) return undefined;
  let splitInlineIndex = -1;
  let splitTextIndex = -1;
  inlines.forEach((inline, inlineIndex) => {
    if (inline.type !== 'text') return;
    const textIndex = inline.text.lastIndexOf('\t');
    if (textIndex >= 0) {
      splitInlineIndex = inlineIndex;
      splitTextIndex = textIndex;
    }
  });
  if (splitInlineIndex < 0) return undefined;

  const splitInline = inlines[splitInlineIndex];
  if (splitInline.type !== 'text') return undefined;
  const left = inlines.slice(0, splitInlineIndex);
  const right = inlines.slice(splitInlineIndex + 1);
  if (splitTextIndex > 0) {
    left.push({
      ...splitInline,
      text: splitInline.text.slice(0, splitTextIndex),
    });
  }
  const pageNumber = splitInline.text.slice(splitTextIndex + 1);
  if (pageNumber) right.unshift({ ...splitInline, text: pageNumber });
  return { left, right };
}

/** 渲染DOC段落内容块。 */
function DocParagraphBlockComponent({ block }: DocParagraphBlockProps) {
  const isTitle = block.role === 'title';
  const isHeading = block.role === 'heading';
  const paragraphStyle = useMemo<CSSProperties>(() => {
    const sourceStyle = docTextStyleToCss(block.style);
    if (block.style?.borderStyle || block.style?.borderWidth) {
      // Word 段落边框贴正文版心，左右缩进只影响边框内文字，不能再次收窄边框外框。
      sourceStyle.marginLeft = 0;
      sourceStyle.marginRight = 0;
    }
    return {
      marginBottom: isTitle ? 18 : isHeading ? 14 : 12,
      fontSize: isTitle ? 22 : isHeading ? 16 : 14,
      lineHeight: isTitle ? 1.45 : isHeading ? 1.65 : 1.8,
      fontWeight: isTitle || isHeading ? 700 : 400,
      ...sourceStyle,
    };
  }, [block.style, isHeading, isTitle]);
  const tocInlines = block.isTableOfContents
    ? splitTableOfContentsInlines(block.inlines)
    : undefined;

  return (
    <p
      className={`office-file-doc-paragraph${
        tocInlines ? ' office-file-doc-paragraph--toc' : ''
      }`}
      style={paragraphStyle}
      data-office-word-outline-target={
        block.outlineLevel !== undefined ? block.id : undefined
      }
    >
      {tocInlines ? (
        <>
          <span className="office-file-doc-paragraph__toc-title">
            <DocInlineContent
              inlines={tocInlines.left}
              fallback=""
              sourceId={`${block.id}-toc-left`}
            />
          </span>
          <span
            className="office-file-doc-paragraph__toc-leader"
            data-office-doc-tab-leader="dot"
            aria-hidden="true"
          />
          <span className="office-file-doc-paragraph__toc-page">
            <DocInlineContent
              inlines={tocInlines.right}
              fallback=""
              sourceId={`${block.id}-toc-right`}
            />
          </span>
        </>
      ) : (
        <DocInlineContent
          inlines={block.inlines}
          fallback={block.text}
          sourceId={block.id}
          preserveBlockTypography={isTitle || isHeading}
        />
      )}
    </p>
  );
}

export const DocParagraphBlock = memo(DocParagraphBlockComponent);
