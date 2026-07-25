import React, { memo, useMemo } from 'react';
import type { DocxDocument } from '../../services/docx/types';
import { OfficeEmpty } from '../../shell/Empty';
import { DocxBlockRenderer } from './DocxBlockRenderer';
import { DocxPageFrame } from './DocxPageFrame';
import './index.less';

/** 定义 DocxViewer 组件可接收的属性。 */
type DocxViewerProps = {
  /** DocxViewerProps 当前关联的标准化文档模型。 */
  document?: DocxDocument;
  /** 当前预览缩放比例。 */
  zoom: number;
};

// DocxViewer 负责 DOCX 页面内容的缩放渲染和滚动布局。
/** 渲染 DocxViewerComponent 组件。 */
function DocxViewerComponent({ document, zoom }: DocxViewerProps) {
  const pages = useMemo(
    () =>
      document
        ? document.pages?.length
          ? document.pages
          : [
              {
                id: 'docx-page-1',
                page: document.page,
                blocks: document.blocks,
              },
            ]
        : [],
    [document],
  );
  if (!document?.blocks.length || !pages.length) {
    return <OfficeEmpty kind="docx" />;
  }

  return (
    <div className="office-file-docx-viewer">
      <div className="office-file-docx-viewer__scroller">
        {pages.map((pageItem) => {
          const contentWidth =
            pageItem.page.width -
            pageItem.page.marginLeft -
            pageItem.page.marginRight;
          return (
            <DocxPageFrame key={pageItem.id} page={pageItem.page} zoom={zoom}>
              {pageItem.blocks.map((block) => (
                <DocxBlockRenderer
                  key={block.id}
                  block={block}
                  availableWidth={contentWidth}
                />
              ))}
            </DocxPageFrame>
          );
        })}
      </div>
    </div>
  );
}

export const DocxViewer = memo(DocxViewerComponent);
