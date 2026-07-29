import React, { memo, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type {
  DocxDocument,
  DocxPageContent,
  DocxPageRegionVariants,
} from '../../services/docx/types';
import { OfficeEmpty } from '../../shell/Empty';
import { WordOutlineSidebar } from '../word-outline/WordOutlineSidebar';
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

/** 按物理页序号选择首页、偶数页或默认页眉页脚。 */
function selectPageRegion<T>(
  variants: DocxPageRegionVariants<T> | undefined,
  pageIndex: number,
  differentFirstPage: boolean | undefined,
  differentEvenOdd: boolean,
): T | undefined {
  if (!variants) return undefined;
  if (pageIndex === 0 && differentFirstPage)
    return variants.first as T | undefined;
  const displayedPageNumber = pageIndex + (differentFirstPage ? 0 : 1);
  if (
    differentEvenOdd &&
    displayedPageNumber % 2 === 0 &&
    variants.even !== undefined
  )
    return variants.even as T;
  return variants.default as T | undefined;
}

/** 根据浏览器实际排版高度为单节流式 DOCX 补充分页。 */
function useMeasuredDocxPages(
  sourcePages: DocxPageContent[],
  preserveSectionPagination: boolean,
) {
  const measureRef = useRef<HTMLDivElement>(null);
  const [measuredPages, setMeasuredPages] = useState<DocxPageContent[]>();

  useLayoutEffect(() => {
    setMeasuredPages(undefined);
    if (preserveSectionPagination) return;
    const articles = Array.from(
      measureRef.current?.querySelectorAll<HTMLElement>(
        '.office-file-docx-page-frame__article',
      ) ?? [],
    );
    if (articles.length !== sourcePages.length) return;

    const measured: DocxPageContent[] = [];
    let didSplit = false;
    sourcePages.forEach((sourcePage, sourcePageIndex) => {
      const elements = Array.from(
        articles[sourcePageIndex].children,
      ) as HTMLElement[];
      if (elements.length !== sourcePage.blocks.length) return;

      const contentHeight =
        sourcePage.page.minHeight -
        sourcePage.page.marginTop -
        sourcePage.page.marginBottom;
      const firstElement = elements[0];
      const lastElement = elements[elements.length - 1];
      const measuredContentHeight =
        firstElement && lastElement
          ? lastElement.offsetTop +
            lastElement.offsetHeight -
            firstElement.offsetTop
          : 0;
      if (measuredContentHeight <= contentHeight + 120) {
        // Word/WPS 允许正文进入页脚预留区；小幅超出不应覆盖源文件保存的分页结果。
        measured.push(sourcePage);
        return;
      }

      let currentBlocks: DocxPageContent['blocks'] = [];
      let currentHeight = 0;
      let flowIndex = 0;
      const pushPage = () => {
        if (!currentBlocks.length) return;
        flowIndex += 1;
        measured.push({
          ...sourcePage,
          id: `${sourcePage.id}-flow-${flowIndex}`,
          blocks: currentBlocks,
        });
        currentBlocks = [];
        currentHeight = 0;
      };
      const appendBlock = (
        block: DocxPageContent['blocks'][number],
        height: number,
      ) => {
        if (
          currentBlocks.length &&
          currentHeight + height > contentHeight + 1
        ) {
          pushPage();
          didSplit = true;
        }
        currentBlocks.push(block);
        currentHeight += height;
      };

      sourcePage.blocks.forEach((block, blockIndex) => {
        const element = elements[blockIndex];
        const nextElement = elements[blockIndex + 1];
        const blockHeight = nextElement
          ? nextElement.offsetTop - element.offsetTop
          : element.offsetHeight +
            Number.parseFloat(
              window.getComputedStyle(element).marginBottom || '0',
            );
        if (block.type !== 'table') {
          appendBlock(block, blockHeight);
          return;
        }
        if (blockHeight <= contentHeight * 0.6) {
          // 可完整放入一页的中小表格由 Word 整体换页，避免只在上一页留下少量表头。
          appendBlock(block, blockHeight);
          return;
        }

        const rows = Array.from(
          element.querySelectorAll<HTMLElement>('tbody > tr'),
        );
        if (rows.length !== block.rows.length) {
          appendBlock(block, blockHeight);
          return;
        }

        let rowStart = 0;
        let rowHeight = 0;
        const appendTableRows = (rowEnd: number) => {
          if (rowEnd <= rowStart) return;
          const tablePart = {
            ...block,
            id: `${block.id}-rows-${rowStart + 1}-${rowEnd}`,
            rows: block.rows.slice(rowStart, rowEnd),
          };
          appendBlock(tablePart, rowHeight);
          rowStart = rowEnd;
          rowHeight = 0;
        };

        rows.forEach((row, rowIndex) => {
          const height = row.getBoundingClientRect().height;
          if (
            rowIndex > rowStart &&
            currentHeight + rowHeight + height > contentHeight + 1
          ) {
            appendTableRows(rowIndex);
            pushPage();
            didSplit = true;
          }
          rowHeight += height;
        });
        appendTableRows(rows.length);
      });
      pushPage();
    });
    if (didSplit) setMeasuredPages(measured);
  }, [preserveSectionPagination, sourcePages]);

  return { measureRef, pages: measuredPages ?? sourcePages };
}

// DocxViewer 负责 DOCX 页面内容的缩放渲染和滚动布局。
/** 渲染 DocxViewerComponent 组件。 */
function DocxViewerComponent({ document, zoom }: DocxViewerProps) {
  const scrollContainerRef = useRef<HTMLDivElement>(null);
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
  const preserveSectionPagination = Boolean(
    document?.preserveSectionPagination,
  );
  const { measureRef, pages } = useMeasuredDocxPages(
    sourcePages,
    preserveSectionPagination,
  );
  const layoutKey = useMemo(
    () => `${zoom}:${pages.map((item) => item.id).join('|')}`,
    [pages, zoom],
  );
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
      {!preserveSectionPagination ? (
        <div
          ref={measureRef}
          className="office-file-docx-viewer__measure"
          aria-hidden="true"
        >
          {sourcePages.map((pageItem) => (
            <DocxPageFrame key={pageItem.id} page={pageItem.page} zoom={100}>
              {renderPageBlocks(pageItem)}
            </DocxPageFrame>
          ))}
        </div>
      ) : null}
      <div className="office-file-docx-viewer__body">
        <WordOutlineSidebar
          items={document.outline ?? []}
          scrollContainerRef={scrollContainerRef}
          documentIdentity={document}
          layoutKey={layoutKey}
        />
        <div
          ref={scrollContainerRef}
          className="office-file-docx-viewer__scroller"
        >
          {pages.map((pageItem, pageIndex) => {
            const differentEvenOdd = Boolean(
              pageItem.headers?.even !== undefined ||
                pageItem.footerPageNumbers?.even !== undefined,
            );
            const firstBodyText = pageItem.blocks.find(
              (block) => block.type === 'paragraph' && block.text,
            );
            // 目录首页在源文档中关闭页眉，后续目录续页恢复默认页眉。
            const suppressHeader =
              firstBodyText?.type === 'paragraph' &&
              firstBodyText.text === '目录';
            const headerBlocks = suppressHeader
              ? undefined
              : selectPageRegion<DocxPageContent['blocks']>(
                  pageItem.headers,
                  pageIndex,
                  pageItem.differentFirstPage,
                  differentEvenOdd,
                );
            const footerPageNumber = selectPageRegion<boolean>(
              pageItem.footerPageNumbers,
              pageIndex,
              pageItem.differentFirstPage,
              differentEvenOdd,
            );
            return (
              <DocxPageFrame
                key={pageItem.id}
                page={pageItem.page}
                zoom={zoom}
                header={
                  headerBlocks?.length
                    ? renderPageBlocks({ ...pageItem, blocks: headerBlocks })
                    : undefined
                }
                footer={
                  footerPageNumber && pageIndex > 0 ? (
                    <span className="office-file-docx-page-frame__page-number">
                      - {pageIndex} -
                    </span>
                  ) : undefined
                }
              >
                {renderPageBlocks(pageItem)}
              </DocxPageFrame>
            );
          })}
        </div>
      </div>
    </div>
  );
}

export const DocxViewer = memo(DocxViewerComponent);
