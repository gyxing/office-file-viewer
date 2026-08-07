import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type { DocImage } from '../../services/doc/types';

/** DOC图片集组件属性。 */
type DocImageGalleryProps = {
  /** 当前文档或页面包含的图片资源。 */
  images: DocImage[];
};

/** 渲染 DOC 文档中的独立图片集。 */
function DocImageGalleryComponent({ images }: DocImageGalleryProps) {
  const messages = useOfficeFileViewerMessages();
  if (!images.length) return null;

  return (
    <section className="office-file-doc-image-gallery">
      <strong className="office-file-doc-image-gallery__title">
        {messages.document.images}
      </strong>
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
              loading="lazy"
              decoding="async"
            />
          </figure>
        ))}
      </div>
    </section>
  );
}

export const DocImageGallery = memo(DocImageGalleryComponent);
