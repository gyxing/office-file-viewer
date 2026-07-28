import React, { memo, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { DocxDocument, DocxPageContent } from '../../services/docx/types';
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

/** 根据浏览器实际排版高度为单节流式 DOCX 补充分页。 */
function useMeasuredDocxPages(sourcePages: DocxPageContent[]) {
  const measureRef = useRef<HTMLDivElement>(null);
  const [measuredPages, setMeasuredPages] = useState<DocxPageContent[]>();

  useLayoutEffect(() => {
    setMeasuredPages(undefined);
    if (sourcePages.length !== 1) return;

    const sourcePage = sourcePages[0];
    const article = measureRef.current?.querySelector<HTMLElement>(
      '.office-file-docx-page-frame__article',
    );
    const elements = Array.from(article?.children ?? []) as HTMLElement[];
    if (elements.length !== sourcePage.blocks.length) return;

    const contentHeight =
      sourcePage.page.minHeight -
      sourcePage.page.marginTop -
      sourcePage.page.marginBottom;
    const blockHeights = elements.map((element, index) => {
      const nextElement = elements[index + 1];
      if (nextElement) return nextElement.offsetTop - element.offsetTop;
      const marginBottom = Number.parseFloat(
        window.getComputedStyle(element).marginBottom || '0',
      );
      return (
        element.offsetHeight +
        (Number.isFinite(marginBottom) ? marginBottom : 0)
      );
    });
    if (
      blockHeights.reduce((sum, height) => sum + height, 0) <=
      contentHeight + 1
    )
      return;

    const groups: DocxPageContent[] = [];
    let currentBlocks: DocxPageContent['blocks'] = [];
    let currentHeight = 0;
    const pushPage = () => {
      if (!currentBlocks.length) return;
      groups.push({
        ...sourcePage,
        id: `${sourcePage.id}-flow-${groups.length + 1}`,
        blocks: currentBlocks,
      });
      currentBlocks = [];
      currentHeight = 0;
    };

    sourcePage.blocks.forEach((block, index) => {
      const blockHeight = blockHeights[index];
      if (
        currentBlocks.length &&
        currentHeight + blockHeight > contentHeight + 1
      )
        pushPage();
      currentBlocks.push(block);
      currentHeight += blockHeight;
    });
    pushPage();
    setMeasuredPages(groups);
  }, [sourcePages]);

  return { measureRef, pages: measuredPages ?? sourcePages };
}

// DocxViewer 负责 DOCX 页面内容的缩放渲染和滚动布局。
/** 渲染 DocxViewerComponent 组件。 */
function DocxViewerComponent({ document, zoom }: DocxViewerProps) {
  const sourcePages = useMemo<DocxPageContent[]>(
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
  const { measureRef, pages } = useMeasuredDocxPages(sourcePages);
  if (!document?.blocks.length || !pages.length) {
    return <OfficeEmpty kind="docx" />;
  }

  const renderPageBlocks = (pageItem: DocxPageContent) => {
    const contentWidth =
      pageItem.page.width -
      pageItem.page.marginLeft -
      pageItem.page.marginRight;
    return pageItem.blocks.map((block) => (
      <DocxBlockRenderer
        key={block.id}
        block={block}
        availableWidth={contentWidth}
      />
    ));
  };

  return (
    <div className="office-file-docx-viewer">
      {sourcePages.length === 1 ? (
        <div
          ref={measureRef}
          className="office-file-docx-viewer__measure"
          aria-hidden="true"
        >
          <DocxPageFrame page={sourcePages[0].page} zoom={100}>
            {renderPageBlocks(sourcePages[0])}
          </DocxPageFrame>
        </div>
      ) : null}
      <div className="office-file-docx-viewer__scroller">
        {pages.map((pageItem) => (
          <DocxPageFrame key={pageItem.id} page={pageItem.page} zoom={zoom}>
            {renderPageBlocks(pageItem)}
          </DocxPageFrame>
        ))}
      </div>
    </div>
  );
}

export const DocxViewer = memo(DocxViewerComponent);
