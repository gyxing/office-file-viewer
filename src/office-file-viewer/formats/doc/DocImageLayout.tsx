// DocImageLayout 渲染 DOC 中连续图片段落形成的图片组，并按内容宽度决定排布。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { DocImage } from '../../services/doc/types';
import {
  DOC_IMAGE_ROW_GAP,
  getDocImageRenderWidth,
  imageRows,
} from './docRenderUtils';

/** 定义 DocImageLayout 组件可接收的属性。 */
type DocImageLayoutProps = {
  /** DocImageLayoutProps 包含的 images 有序集合。 */
  images: DocImage[];
  /** DocImageLayoutProps 的 contentWidth 尺寸或坐标，单位为标准化渲染像素。 */
  contentWidth: number;
};

/** 渲染 DocImageLayoutComponent 组件。 */
function DocImageLayoutComponent({
  images,
  contentWidth,
}: DocImageLayoutProps) {
  const rows = useMemo(
    () => imageRows(images, contentWidth),
    [contentWidth, images],
  );

  if (!images.length) return null;

  return (
    <div className="office-file-doc-image-layout">
      {rows.map((row) => (
        <div
          key={row.map((image) => image.id).join('-')}
          className="office-file-doc-image-layout__row"
          style={{ maxWidth: contentWidth }}
        >
          {row.map((image) => {
            const rowWidth = getDocImageRenderWidth(
              image,
              contentWidth,
              row.length,
            );
            const aspectRatio =
              image.width && image.height ? image.width / image.height : 0;
            const centerSingleImage =
              row.length === 1 && aspectRatio >= 0.75 && aspectRatio <= 1.25;
            const figureStyle: CSSProperties = {
              width: rowWidth,
              maxWidth:
                row.length > 1
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
                <img
                  className="office-file-doc-image-layout__img"
                  src={image.src}
                  alt={image.caption ?? image.id}
                />
              </figure>
            );
          })}
        </div>
      ))}
    </div>
  );
}

export const DocImageLayout = memo(DocImageLayoutComponent);
