// DocImageLayout 渲染 DOC 中连续图片段落形成的图片组，并按内容宽度决定排布。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import {
  DOC_IMAGE_ROW_GAP,
  getDocImageRenderWidth,
  imageRows,
} from '../../services/doc/docPagination';
import type { DocImage, DocTextStyle } from '../../services/doc/types';
import { OfficePreviewableImage } from '../../shared/image-preview';

/** DOC图片布局组件属性。 */
type DocImageLayoutProps = {
  /** 当前文档或页面包含的图片资源。 */
  images: DocImage[];
  /** 可用于排版内容的宽度，单位为标准化渲染像素。 */
  contentWidth: number;
  /** 图片所在源段落的水平对齐方式。 */
  alignment?: DocTextStyle['textAlign'];
  /** 图片段落前由源文档空段落或显式间距形成的留白。 */
  spacingBefore?: number;
};

/** 渲染DOC图片布局。 */
function DocImageLayoutComponent({
  images,
  contentWidth,
  alignment,
  spacingBefore,
}: DocImageLayoutProps) {
  const rows = useMemo(
    () => imageRows(images, contentWidth),
    [contentWidth, images],
  );
  const pageCanvas = images.length === 1 ? images[0].pageInsets : undefined;

  if (!images.length) return null;

  return (
    <div
      className="office-file-doc-image-layout"
      style={
        pageCanvas
          ? {
              marginTop: -pageCanvas.top,
              marginRight: -pageCanvas.right,
              marginBottom: -pageCanvas.bottom,
              marginLeft: -pageCanvas.left,
            }
          : { marginTop: spacingBefore }
      }
    >
      {rows.map((row) => {
        const firstImage = row[0];
        const firstAspectRatio =
          firstImage?.width && firstImage.height
            ? firstImage.width / firstImage.height
            : 0;
        const centerSingleImage =
          !pageCanvas &&
          row.length === 1 &&
          (alignment === 'center' ||
            (!alignment &&
              firstAspectRatio >= 0.75 &&
              firstAspectRatio <= 1.25));

        return (
          <div
            key={row.map((image) => image.id).join('-')}
            className="office-file-doc-image-layout__row"
            style={{
              maxWidth: pageCanvas ? firstImage.width : contentWidth,
              justifyContent: pageCanvas
                ? 'flex-start'
                : alignment === 'center' || centerSingleImage
                ? 'center'
                : alignment === 'right'
                ? 'flex-end'
                : 'flex-start',
            }}
          >
            {row.map((image) => {
              const rowWidth = getDocImageRenderWidth(
                image,
                contentWidth,
                row.length,
              );
              const figureStyle: CSSProperties = {
                width: rowWidth,
                maxWidth: pageCanvas
                  ? 'none'
                  : row.length > 1
                  ? `calc((100% - ${DOC_IMAGE_ROW_GAP}px) / ${row.length})`
                  : '100%',
                marginInline: centerSingleImage ? 'auto' : undefined,
              };

              return (
                <figure
                  key={image.id}
                  className="office-file-doc-image-layout__figure"
                  style={figureStyle}
                >
                  <OfficePreviewableImage
                    previewId={image.id}
                    previewName={image.caption}
                    previewMimeType={image.mimeType}
                    previewSource={image.src}
                    className="office-file-doc-image-layout__img"
                    src={image.src}
                    alt={image.caption ?? image.id}
                    loading="lazy"
                    decoding="async"
                  />
                </figure>
              );
            })}
          </div>
        );
      })}
    </div>
  );
}

export const DocImageLayout = memo(DocImageLayoutComponent);
