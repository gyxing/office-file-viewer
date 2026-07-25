// DocInlineContent 渲染 DOC 段落或表格单元格内的文本和图片片段。
import React, { memo } from 'react';
import type { DocTextInline } from '../../services/doc/types';
import { inlineStyleToCss } from './docRenderUtils';

/** 定义 DocInlineContent 组件可接收的属性。 */
type DocInlineContentProps = {
  /** DocInlineContentProps 包含的 inlines 有序集合。 */
  inlines?: DocTextInline[];
  /** DocInlineContentProps 的 fallback 文本值。 */
  fallback: string;
  /** 是否保留块级模型自身的字体与段落样式；未提供时使用来源格式或渲染器的默认行为。 */
  preserveBlockTypography?: boolean;
};

/** 渲染 DocInlineContentComponent 组件。 */
function DocInlineContentComponent({
  inlines,
  fallback,
  preserveBlockTypography,
}: DocInlineContentProps) {
  if (!inlines?.length) return <>{fallback}</>;

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
            {inline.text}
          </span>
        ),
      )}
    </>
  );
}

export const DocInlineContent = memo(DocInlineContentComponent);
