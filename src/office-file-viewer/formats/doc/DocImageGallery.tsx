// DocImageGallery 展示未能稳定锚定到正文位置的 DOC 图片。
import { Typography } from 'antd';
import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type { DocImage } from '../../services/doc/types';

/** 定义 DocImageGallery 组件可接收的属性。 */
type DocImageGalleryProps = {
  /** DocImageGalleryProps 包含的 images 有序集合。 */
  images: DocImage[];
};

/** 渲染 DocImageGalleryComponent 组件。 */
function DocImageGalleryComponent({ images }: DocImageGalleryProps) {
  const messages = useOfficeFileViewerMessages();
  if (!images.length) return null;

  return (
    <section className="office-file-doc-image-gallery">
      <Typography.Text strong className="office-file-doc-image-gallery__title">
        {messages.document.images}
      </Typography.Text>
      <div className="office-file-doc-image-gallery__grid">
        {images.map((image) => (
          <figure
            key={image.id}
            className="office-file-doc-image-gallery__figure"
          >
            <img
              className="office-file-doc-image-gallery__img"
              src={image.src}
              alt={image.caption ?? image.id}
            />
          </figure>
        ))}
      </div>
    </section>
  );
}

export const DocImageGallery = memo(DocImageGalleryComponent);
