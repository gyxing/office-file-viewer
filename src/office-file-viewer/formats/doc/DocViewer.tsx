import React, { memo, useMemo } from 'react';
import type { DocDocument } from '../../services/doc/types';
import { OfficeEmpty } from '../../shell/Empty';
import { DocContentRenderer } from './DocContentRenderer';
import { DocImageGallery } from './DocImageGallery';
import { DocPageFrame } from './DocPageFrame';
import { paginateDocBlocks } from './docRenderUtils';
import './index.less';

/** 定义 DocViewer 组件可接收的属性。 */
type DocViewerProps = {
  /** DocViewerProps 当前关联的标准化文档模型。 */
  document?: DocDocument;
  /** 当前预览缩放比例。 */
  zoom: number;
};

/** 提取并汇总 `collectAnchoredImageIds` 返回的数据。 */
function collectAnchoredImageIds(document?: DocDocument) {
  const ids = new Set<string>();
  if (document?.headerImage) ids.add(document.headerImage.id);
  document?.blocks.forEach((block) => {
    if (block.type === 'paragraph') {
      block.inlines?.forEach((inline) => {
        if (inline.type === 'image') ids.add(inline.image.id);
      });
    } else if (block.type === 'table') {
      block.rows.forEach((row) =>
        row.cells.forEach((cell) =>
          cell.inlines?.forEach((inline) => {
            if (inline.type === 'image') ids.add(inline.image.id);
          }),
        ),
      );
    } else {
      block.items.forEach((item) =>
        item.inlines?.forEach((inline) => {
          if (inline.type === 'image') ids.add(inline.image.id);
        }),
      );
    }
  });
  return ids;
}

// DocViewer 负责旧版 DOC/WPS 降级预览的固定警告和页面滚动区。
/** 渲染 DocViewerComponent 组件。 */
function DocViewerComponent({ document, zoom }: DocViewerProps) {
  const page = document?.page;
  const contentWidth = page
    ? page.width - page.marginLeft - page.marginRight
    : 0;
  const pages = useMemo(
    () =>
      document && page
        ? paginateDocBlocks(document.blocks, page, contentWidth)
        : [],
    [contentWidth, document, page],
  );
  const anchoredImageIds = useMemo(
    () => collectAnchoredImageIds(document),
    [document],
  );
  const unanchoredImages = useMemo(
    () =>
      document?.images.filter((image) => !anchoredImageIds.has(image.id)) ?? [],
    [anchoredImageIds, document],
  );

  if (!document?.blocks.length || !page || !pages.length) {
    return <OfficeEmpty kind="doc" />;
  }

  return (
    <div className="office-file-doc-viewer">
      {document.warnings.length ? (
        <div className="office-file-doc-viewer__notice" role="alert">
          {document.warnings.join(' ')}
        </div>
      ) : null}
      <div className="office-file-doc-viewer__scroller">
        {pages.map((docPage, pageIndex) => (
          <DocPageFrame
            key={docPage.id}
            page={page}
            zoom={zoom}
            headerImage={
              docPage.blocks.length > 0 &&
              docPage.blocks.every(
                (block) =>
                  block.type === 'paragraph' &&
                  /^\s*\d+(?:\.\d+)*\s+.+\s+-\s*.+\s*-\s*$/.test(block.text),
              )
                ? undefined
                : document.headerImage
            }
            footerText={
              document.footerPageNumbers && pageIndex > 0
                ? `- ${pageIndex} -`
                : undefined
            }
          >
            <DocContentRenderer
              blocks={docPage.blocks}
              contentWidth={contentWidth}
            />
            {pageIndex === pages.length - 1 ? (
              <DocImageGallery images={unanchoredImages} />
            ) : null}
          </DocPageFrame>
        ))}
      </div>
    </div>
  );
}

export const DocViewer = memo(DocViewerComponent);
