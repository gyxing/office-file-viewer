// DocInlineContent 渲染 DOC 段落或表格单元格内的文本和图片片段。
import React, { memo } from 'react';
import type {
  DocBookmarkInline,
  DocImageInline,
  DocTextInline,
  DocTextRunInline,
} from '../../services/doc/types';
import { useOfficeHyperlink } from '../../shared/hyperlink';
import { OfficePreviewableImage } from '../../shared/image-preview';
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
  /** 行内内容所属块或单元格的稳定标识。 */
  sourceId?: string;
};

/** 恢复 Word 表格中斜杠与后续数字不拆行的排版语义。 */
function withWordTableLineBreaks(text: string) {
  return text.replace(/\/(?=\d)/g, '\u200b/\u2060');
}

/** 渲染一个可选链接的 DOC 文字片段。 */
function DocTextRunContent({
  inline,
  sourceId,
  preserveBlockTypography,
  wordTableLineBreaks,
}: {
  /** 当前文字片段。 */
  inline: DocTextRunInline;
  /** 当前片段的稳定标识。 */
  sourceId: string;
  /** 是否保留块级字体样式。 */
  preserveBlockTypography?: boolean;
  /** 是否应用表格数字分隔符规则。 */
  wordTableLineBreaks?: boolean;
}) {
  const hyperlinkProps = useOfficeHyperlink<HTMLSpanElement>({
    hyperlink: inline.hyperlink,
    source: { type: 'text', id: sourceId },
  });
  return (
    <span
      {...hyperlinkProps}
      style={inlineStyleToCss(inline.style, { preserveBlockTypography })}
    >
      {wordTableLineBreaks ? withWordTableLineBreaks(inline.text) : inline.text}
    </span>
  );
}

/** 渲染一个可选链接且仍可双击预览的 DOC 图片片段。 */
function DocImageInlineContent({
  inline,
  sourceId,
}: {
  /** 当前图片片段。 */
  inline: DocImageInline;
  /** 当前片段的稳定标识。 */
  sourceId: string;
}) {
  const hyperlinkProps = useOfficeHyperlink<HTMLImageElement>({
    hyperlink: inline.image.hyperlink,
    source: { type: 'image', id: sourceId },
  });
  if (inline.image.pageDrawingLayer) return null;
  return (
    <span className="office-file-doc-inline-image">
      <OfficePreviewableImage
        {...hyperlinkProps}
        previewId={inline.image.id}
        previewName={inline.image.caption}
        previewMimeType={inline.image.mimeType}
        previewSource={inline.image.src}
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
  );
}

/** 渲染 DOC/WPS 内部跳转使用的零宽书签标记。 */
function DocBookmarkInlineContent({ inline }: { inline: DocBookmarkInline }) {
  return (
    <span
      className="office-file-word-bookmark"
      data-office-word-bookmark={inline.name}
      data-office-word-bookmark-id={inline.markerId}
      aria-hidden="true"
    />
  );
}

/** 渲染 DOC 段落中的行内文字和图片。 */
function DocInlineContentComponent({
  inlines,
  fallback,
  preserveBlockTypography,
  wordTableLineBreaks,
  sourceId = 'doc-inline',
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
          <DocImageInlineContent
            key={`${inline.image.id}-${index}`}
            inline={inline}
            sourceId={`${sourceId}-image-${inline.image.id}`}
          />
        ) : inline.type === 'bookmark' ? (
          <DocBookmarkInlineContent
            key={`${inline.markerId}-${index}`}
            inline={inline}
          />
        ) : (
          <DocTextRunContent
            key={`${inline.text}-${index}`}
            inline={inline}
            sourceId={`${sourceId}-text-${index}`}
            preserveBlockTypography={preserveBlockTypography}
            wordTableLineBreaks={wordTableLineBreaks}
          />
        ),
      )}
    </>
  );
}

export const DocInlineContent = memo(DocInlineContentComponent);
