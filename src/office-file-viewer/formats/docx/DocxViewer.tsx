import React, {
  lazy,
  memo,
  Suspense,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import type {
  DocxPageContent,
  DocxPageRegionVariants,
} from '../../services/docx/types';
import type { OfficeFileViewerPreviewState } from '../../services/parsing/internalTypes';
import { collectWordPerformanceStats } from '../../services/word/collectWordPerformanceStats';
import { createMaterializedWordPageSource } from '../../services/word/createMaterializedWordPageSource';
import { createMemoryWordOutlineProvider } from '../../services/word/createMemoryWordOutlineProvider';
import { useExternalStoreSnapshot } from '../../shared/react/useExternalStoreSnapshot';
import { OfficeEmpty } from '../../shell/Empty';
import { useWordOutlinePresence } from '../word-outline/useWordOutlinePresence';
import { WordOutlineSidebar } from '../word-outline/WordOutlineSidebar';
import type { WordPageNavigationController } from '../word-pages/types';
import { VirtualWordPageList } from '../word-pages/VirtualWordPageList';
import { WordBlockPageIndex } from '../word-pages/WordBlockPageIndex';
import { useWordPerformanceProfile } from '../word-performance/useWordPerformanceProfile';
import { DocxBlockRenderer } from './DocxBlockRenderer';
import { DocxCharacterSpacingContext } from './DocxInlineContent';
import { DocxPageFrame } from './DocxPageFrame';
import {
  paginateMeasuredDocxPage,
  type DocxMeasuredBlock,
  type DocxMeasurementBatch,
} from './docxPagination';
import './index.less';
import { measureDocxParagraphLines } from './measureDocxParagraphLines';

const LazyDocxMeasureHost = lazy(() =>
  import('./DocxMeasureHost').then((module) => ({
    default: module.DocxMeasureHost,
  })),
);

/** DOCX Viewer 可以消费的物化或按需预览。 */
type DocxPreview = Extract<
  OfficeFileViewerPreviewState,
  { previewKind: 'docx' }
>;

/** DOCX预览器组件属性。 */
type DocxViewerProps = {
  /** 当前 DOCX 的物化或按需预览。 */
  preview: DocxPreview;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 文档大纲当前是否展开。 */
  showOutline: boolean;
  /** 关闭文档大纲。 */
  onCloseOutline: () => void;
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
  reportPaginationDuration: (durationMs: number) => void,
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
    const startedAt = performance.now();

    const measured: DocxPageContent[] = [];
    sourcePages.forEach((sourcePage, sourcePageIndex) => {
      const elements = Array.from(
        articles[sourcePageIndex].children,
      ) as HTMLElement[];
      if (elements.length !== sourcePage.blocks.length) return;
      const measurements: DocxMeasuredBlock[] = sourcePage.blocks.map(
        (block, blockIndex) => {
          const element = elements[blockIndex];
          const nextElement = elements[blockIndex + 1];
          const blockHeight = nextElement
            ? nextElement.offsetTop - element.offsetTop
            : element.offsetHeight +
              Number.parseFloat(
                window.getComputedStyle(element).marginBottom || '0',
              );
          const rowHeights =
            block.type === 'table'
              ? Array.from(
                  element.querySelectorAll<HTMLElement>('tbody > tr'),
                ).map((row) => row.getBoundingClientRect().height)
              : undefined;
          return {
            block,
            height: blockHeight,
            rowHeights,
            originalTableRowCount:
              block.type === 'table' ? block.rows.length : undefined,
            ...measureDocxParagraphLines(element, block, blockHeight),
          };
        },
      );
      measured.push(...paginateMeasuredDocxPage(sourcePage, measurements));
    });
    const didSplit =
      measured.length !== sourcePages.length ||
      measured.some((page, index) => page !== sourcePages[index]);
    if (didSplit) setMeasuredPages(measured);
    reportPaginationDuration(performance.now() - startedAt);
  }, [preserveSectionPagination, reportPaginationDuration, sourcePages]);

  return { measureRef, pages: measuredPages ?? sourcePages };
}

// DocxViewer 负责 DOCX 页面内容的缩放渲染和滚动布局。
/** 协调 DOCX 页面、大纲和按需内容加载。 */
function DocxViewerComponent({
  preview,
  zoom,
  showOutline,
  onCloseOutline,
}: DocxViewerProps) {
  const document =
    preview.mode === 'materialized' ? preview.model.document : undefined;
  const source = preview.mode === 'source' ? preview.source : undefined;
  const summary = preview.mode === 'source' ? preview.summary : undefined;
  const documentSessionId = preview.sessionId;
  const shouldRenderOutline = useWordOutlinePresence(
    showOutline,
    documentSessionId,
  );
  const scrollContainerRef = useRef<HTMLDivElement>(null);
  const materializedSourcePages = useMemo<DocxPageContent[]>(
    () =>
      document && !source
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
    [document, source],
  );
  const preserveSectionPagination = Boolean(
    source
      ? summary?.preserveSectionPagination
      : document?.preserveSectionPagination,
  );
  const characterSpacingControl = source
    ? summary?.characterSpacingControl
    : document?.characterSpacingControl;
  const compressPunctuation =
    characterSpacingControl === 'compressPunctuation' ||
    characterSpacingControl === 'compressPunctuationAndJapaneseKana';
  const materializedOutlineItems = useMemo(
    () => (source || !shouldRenderOutline ? [] : document?.outline ?? []),
    [document?.outline, shouldRenderOutline, source],
  );
  const materializedOutlineProvider = useMemo(
    () => createMemoryWordOutlineProvider(materializedOutlineItems),
    [materializedOutlineItems],
  );
  const materializedPerformanceStats = useMemo(
    () =>
      document && !source
        ? collectWordPerformanceStats(document, {
            estimatedPageCount: materializedSourcePages.length,
          })
        : {
            estimatedPageCount: 0,
            outlineCount: 0,
            paragraphCount: 0,
            tableRowCount: 0,
            imageCount: 0,
            drawingCount: 0,
            textLength: 0,
            slowPagination: false,
          },
    [document, materializedSourcePages.length, source],
  );
  const { profile: materializedProfile, reportPaginationDuration } =
    useWordPerformanceProfile(documentSessionId, materializedPerformanceStats);
  const { measureRef, pages: materializedPages } = useMeasuredDocxPages(
    materializedSourcePages,
    preserveSectionPagination,
    reportPaginationDuration,
  );
  const materializedPageSource = useMemo(
    () =>
      createMaterializedWordPageSource(materializedPages, {
        getId: (pageItem) => pageItem.id,
        getEstimatedContentHeight: (pageItem) => pageItem.page.minHeight,
        getSourceBlockIds: (pageItem) =>
          pageItem.blocks.flatMap((block) => [
            block.id,
            ...(block.sourceBlockId ? [block.sourceBlockId] : []),
          ]),
      }),
    [materializedPages],
  );
  useEffect(
    () => () => void materializedPageSource.dispose(),
    [materializedPageSource],
  );
  const pageSource = source ?? materializedPageSource;
  const pageSnapshot = useExternalStoreSnapshot(pageSource);
  const outlineItems = useMemo(
    () =>
      shouldRenderOutline
        ? source
          ? source.getOutlineItems()
          : materializedOutlineItems
        : [],
    [
      materializedOutlineItems,
      pageSnapshot.revision,
      shouldRenderOutline,
      source,
    ],
  );
  const outlineProvider = source?.outline ?? materializedOutlineProvider;
  const profile = source?.getPerformanceProfile() ?? materializedProfile;
  const blockPageIndex = useMemo(() => {
    const index = new WordBlockPageIndex();
    pageSnapshot.pages.forEach((meta) => index.replacePage(meta));
    return index;
  }, [pageSnapshot.pages, pageSnapshot.revision]);
  const pageNavigationControllerRef = useRef<WordPageNavigationController>();
  const layoutKey = useMemo(
    () =>
      source
        ? `${documentSessionId}:${zoom}:source`
        : `${zoom}:${materializedPages.map((item) => item.id).join('|')}`,
    [documentSessionId, materializedPages, source, zoom],
  );

  const renderPageBlocks = useCallback((pageItem: DocxPageContent) => {
    const contentWidth =
      pageItem.page.width -
      pageItem.page.marginLeft -
      pageItem.page.marginRight;
    return pageItem.blocks.map((block) => (
      <DocxBlockRenderer
        key={block.id}
        block={block}
        availableWidth={contentWidth}
        maximumWidth={pageItem.page.width}
      />
    ));
  }, []);
  const renderPage = useCallback(
    (pageItem: DocxPageContent, pageIndex: number) => {
      const differentEvenOdd = Boolean(
        pageItem.headers?.even !== undefined ||
          pageItem.footerPageNumbers?.even !== undefined,
      );
      const headerBlocks = selectPageRegion<DocxPageContent['blocks']>(
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
      const displayedPageNumber =
        pageIndex + (pageItem.differentFirstPage ? 0 : 1);
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
            footerPageNumber && displayedPageNumber > 0 ? (
              <span className="office-file-docx-page-frame__page-number">
                - {displayedPageNumber} -
              </span>
            ) : undefined
          }
        >
          {renderPageBlocks(pageItem)}
        </DocxPageFrame>
      );
    },
    [renderPageBlocks, zoom],
  );
  const measurementBatch = source?.getMeasurementBatch();
  const renderMeasurementBlock = useCallback(
    (block: DocxPageContent['blocks'][number]) => {
      const page = source?.getMeasurementBatch()?.sourcePage;
      const availableWidth = page
        ? page.page.width - page.page.marginLeft - page.page.marginRight
        : 0;
      return (
        <DocxBlockRenderer
          key={block.id}
          block={block}
          availableWidth={availableWidth}
          maximumWidth={page?.page.width}
        />
      );
    },
    [source],
  );
  const handleMeasured = useCallback(
    (
      batch: DocxMeasurementBatch,
      blocks: readonly DocxMeasuredBlock[],
      durationMs: number,
    ) => {
      void source
        ?.commitMeasurement(batch, blocks, durationMs)
        .catch((error) => source.failMeasurement(batch, error));
    },
    [source],
  );
  const handleMeasurementError = useCallback(
    (batch: DocxMeasurementBatch, error: unknown) => {
      source?.failMeasurement(batch, error);
    },
    [source],
  );

  if (
    (!source && (!document?.blocks.length || !materializedPages.length)) ||
    (source && !summary)
  ) {
    return <OfficeEmpty kind="docx" />;
  }

  return (
    <DocxCharacterSpacingContext.Provider value={compressPunctuation}>
      <div
        className="office-file-docx-viewer"
        data-word-source-mode={source ? 'progressive' : 'materialized'}
        data-character-spacing-control={characterSpacingControl}
      >
        {source ? (
          <Suspense fallback={null}>
            <LazyDocxMeasureHost
              batch={measurementBatch}
              renderBlock={renderMeasurementBlock}
              onMeasured={handleMeasured}
              onError={handleMeasurementError}
            />
          </Suspense>
        ) : !preserveSectionPagination ? (
          <div
            ref={measureRef}
            className="office-file-docx-viewer__measure"
            aria-hidden="true"
          >
            {materializedSourcePages.map((pageItem) => (
              <DocxPageFrame key={pageItem.id} page={pageItem.page} zoom={100}>
                {renderPageBlocks(pageItem)}
              </DocxPageFrame>
            ))}
          </div>
        ) : null}
        <div className="office-file-docx-viewer__body">
          <WordOutlineSidebar
            visible={showOutline}
            activated={shouldRenderOutline}
            items={outlineItems}
            provider={outlineProvider}
            outlineMode={profile.outlineMode}
            pageMode={profile.pageMode}
            pageSource={pageSource}
            blockPageIndex={blockPageIndex}
            pageNavigationControllerRef={pageNavigationControllerRef}
            scrollContainerRef={scrollContainerRef}
            documentSessionId={documentSessionId}
            layoutKey={layoutKey}
            onClose={onCloseOutline}
          />
          <div
            ref={scrollContainerRef}
            className="office-file-docx-viewer__scroller"
          >
            {profile.pageMode === 'windowed' ? (
              <VirtualWordPageList
                source={pageSource}
                scrollerRef={scrollContainerRef}
                layoutRevision={layoutKey}
                zoom={zoom}
                navigationControllerRef={pageNavigationControllerRef}
                renderPage={renderPage}
              />
            ) : (
              materializedPages.map(renderPage)
            )}
          </div>
        </div>
      </div>
    </DocxCharacterSpacingContext.Provider>
  );
}

export const DocxViewer = memo(DocxViewerComponent);
