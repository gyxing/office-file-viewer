// DocxImage 渲染 DOCX 行内图片，并保留文档解析出的图片宽度。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { DocxImageInline } from '../../services/docx/types';
import { useOfficeResourceUrl } from '../../services/resource-store/useOfficeResourceUrl';
import { calculatePositionStyle } from './positionUtils';

/** DOCX图片组件属性。 */
type DocxImageProps = {
  /** 当前负责渲染的行内内容模型。 */
  inline: DocxImageInline;
};

// 自定义变量把解析后的图片尺寸交给 Less，避免静态样式散落在 JSX 中。
/** DOCX 图片渲染样式。 */
type DocxImageStyle = CSSProperties & {
  /** 传递给样式表的形状宽度 CSS 自定义属性。 */
  '--office-file-docx-inline-image-width': string;
  /** 传递给样式表的形状高度 CSS 自定义属性。 */
  '--office-file-docx-inline-image-height': string;
};

/** 渲染DOCX图片。 */
function DocxImageComponent({ inline }: DocxImageProps) {
  const image = inline.image;
  const positionStyle = calculatePositionStyle(image.position);
  const resource = useOfficeResourceUrl(image.src);

  const imageStyle = useMemo<DocxImageStyle>(
    () => ({
      '--office-file-docx-inline-image-width': `${image.width}px`,
      '--office-file-docx-inline-image-height': `${image.height}px`,
      ...positionStyle,
      maxWidth: image.position ? 'none' : undefined,
    }),
    [image.height, image.position, image.width, positionStyle],
  );

  return (
    <img
      className="office-file-docx-inline-image"
      src={resource.url}
      alt={image.alt ?? ''}
      title={image.name}
      style={imageStyle}
      loading="lazy"
      decoding="async"
      aria-busy={resource.loading || undefined}
      data-resource-error={resource.error ? 'true' : undefined}
    />
  );
}

export const DocxImage = memo(DocxImageComponent);
