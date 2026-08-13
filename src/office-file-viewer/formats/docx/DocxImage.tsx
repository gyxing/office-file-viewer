// DocxImage 渲染 DOCX 行内图片，并保留文档解析出的图片宽度。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { DocxImageInline } from '../../services/docx/types';
import { useOfficeResourceUrl } from '../../services/resource-store/useOfficeResourceUrl';
import { useOfficeHyperlink } from '../../shared/hyperlink';
import { OfficePreviewableImage } from '../../shared/image-preview';
import { calculatePositionStyle } from './positionUtils';
import {
  resolveDocxCroppedImageStyle,
  useDocxImageColorChange,
} from './useDocxImageColorChange';

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
  const effectedResource = useDocxImageColorChange(
    resource.url,
    image.colorChange,
  );
  const hyperlinkProps = useOfficeHyperlink<HTMLImageElement>({
    hyperlink: image.hyperlink,
    source: { type: 'image', id: image.id },
  });

  const imageStyle = useMemo<DocxImageStyle>(
    () => ({
      '--office-file-docx-inline-image-width': `${image.width}px`,
      '--office-file-docx-inline-image-height': `${image.height}px`,
      ...positionStyle,
      maxWidth: image.position ? 'none' : undefined,
    }),
    [image.height, image.position, image.width, positionStyle],
  );
  const croppedImageStyle = useMemo<CSSProperties | undefined>(
    () =>
      image.crop ? resolveDocxCroppedImageStyle(image.crop) : undefined,
    [image.crop],
  );

  const previewableImage = (
    <OfficePreviewableImage
      {...hyperlinkProps}
      className="office-file-docx-inline-image"
      previewId={image.id}
      previewName={image.name}
      previewSource={image.src}
      src={effectedResource.url}
      alt={image.alt ?? ''}
      title={image.name}
      style={croppedImageStyle ?? imageStyle}
      loading="lazy"
      decoding="async"
      aria-busy={resource.loading || effectedResource.loading || undefined}
      data-resource-error={
        resource.error || effectedResource.error ? 'true' : undefined
      }
    />
  );
  if (!croppedImageStyle) return previewableImage;

  return (
    <span className="office-file-docx-inline-image-crop" style={imageStyle}>
      {previewableImage}
    </span>
  );
}

export const DocxImage = memo(DocxImageComponent);
