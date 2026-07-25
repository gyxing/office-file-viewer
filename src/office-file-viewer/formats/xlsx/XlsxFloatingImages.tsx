// XlsxFloatingImages 渲染锚定在工作表画布上的浮动图片。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { XlsxImage } from '../../services/xlsx/types';

/** 定义 XlsxFloatingImages 组件可接收的属性。 */
type XlsxFloatingImagesProps = {
  /** XlsxFloatingImagesProps 包含的 images 有序集合。 */
  images: XlsxImage[];
};

/** 渲染 XlsxFloatingImage 组件。 */
function XlsxFloatingImage({
  image,
}: {
  /** 当前结构 当前关联的图片资源或图片模型。 */ image: XlsxImage;
}) {
  const imageStyle = useMemo<CSSProperties>(
    () => ({
      left: 48 + image.x,
      top: 28 + image.y,
      width: image.width,
      height: image.height,
    }),
    [image.height, image.width, image.x, image.y],
  );

  return (
    <img
      className="office-file-xlsx-sheet-grid__floating-image"
      src={image.src}
      alt={image.alt ?? ''}
      title={image.name}
      style={imageStyle}
      onError={(event) => {
        event.currentTarget.setAttribute(
          'aria-label',
          image.alt ? `${image.alt}（图片加载失败）` : '图片加载失败',
        );
        event.currentTarget.setAttribute('data-load-error', 'true');
      }}
    />
  );
}

const MemoXlsxFloatingImage = memo(XlsxFloatingImage);

/** 渲染 XlsxFloatingImagesComponent 组件。 */
function XlsxFloatingImagesComponent({ images }: XlsxFloatingImagesProps) {
  return (
    <>
      {images.map((image) => (
        <MemoXlsxFloatingImage key={image.id} image={image} />
      ))}
    </>
  );
}

export const XlsxFloatingImages = memo(XlsxFloatingImagesComponent);
