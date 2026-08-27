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
import { useOfficeFileViewerMessages } from '../../locale';
import {
  createMemoryOfficeAnnotationSource,
  type WordRevisionMode,
} from '../../services/annotations';
import {
  paginateMeasuredDocxPage,
  type DocxMeasuredBlock,
  type DocxMeasurementBatch,
} from '../../services/docx/docxPagination';
import {
  projectDocxBlock,
  projectDocxPageContent,
} from '../../services/docx/projectDocxRevisionMode';
import type {
  DocxPageContent,
  DocxPageRegionVariants,
} from '../../services/docx/types';
import type { OfficeFileViewerPreviewState } from '../../services/parsing/internalTypes';
import { collectWordPerformanceStats } from '../../services/word/collectWordPerformanceStats';
import { createMaterializedWordPageSource } from '../../services/word/createMaterializedWordPageSource';
import { createMemoryWordOutlineProvider } from '../../services/word/createMemoryWordOutlineProvider';
import { collectDocxSearchBlocks } from '../../services/word/WordSearchProvider';
import { useOfficeAnnotationSourceRegistration } from '../../shared/annotations';
import { useOfficeEmbeddedFontsReady } from '../../shared/fonts/OfficeFontProvider';
import { OfficeImagePreviewContext } from '../../shared/image-preview/OfficeImagePreviewContext';
import { useExternalStoreSnapshot } from '../../shared/react/useExternalStoreSnapshot';
import { OfficeSpinner } from '../../shared/ui/OfficeSpinner';
import { OfficeWatermarkSurface } from '../../shared/watermark';
import { OfficePreviewEmpty } from '../common/OfficePreviewEmpty';
import { useOfficeSearchProviderRegistration } from '../search/OfficeSearchContext';
import { useWordTargetNavigation } from '../word-hyperlink/useWordTargetNavigation';
import { useWordOutlinePresence } from '../word-outline/useWordOutlinePresence';
import { WordOutlineSidebar } from '../word-outline/WordOutlineSidebar';
import type { WordPageNavigationController } from '../word-pages/types';
import { VirtualWordPageList } from '../word-pages/VirtualWordPageList';
import { WordBlockPageIndex } from '../word-pages/WordBlockPageIndex';
import { useWordPerformanceProfile } from '../word-performance/useWordPerformanceProfile';
import '../word-review/index.less';
import { useWordAnnotationNavigation } from '../word-review/useWordAnnotationNavigation';
import { WORD_MARKUP_BASE_RAIL_WIDTH } from '../word-review/wordMarkupCalloutLayout';
import {
  collectDocxNoteReferences,
  WordNoteBlock,
} from '../word-review/WordNoteBlock';
import { WordReviewOverlay } from '../word-review/WordReviewOverlay';
import { WordRevisionModeContext } from '../word-review/WordRevisionText';
import { useWordSearchNavigation } from '../word-search/useWordSearchNavigation';
import { DocxBlockRenderer } from './DocxBlockRenderer';
import { DocxCharacterSpacingContext } from './DocxInlineContent';
import { DocxPageFrame } from './DocxPageFrame';
import {
  resolveDocxSpacingBefore,
  shouldSuppressDocxContextualSpacing,
} from './docxParagraphSpacing';
import {
  alignDocxInlineObjectParagraphToGrid,
  measureVisibleDocxBlockHeight,
} from './docxRenderUtils';
import './index.less';
import { measureDocxFootnotes } from './measureDocxFootnotes';
import { measureDocxParagraphLines } from './measureDocxParagraphLines';
import { measureDocxTableRows } from './measureDocxTableRows';

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
  /** 当前采用的 Word 修订内容投影模式。 */
  wordRevisionMode: WordRevisionMode;
  /** 关闭文档大纲。 */
  onCloseOutline: () => void;
  /** 搜索能力启用时切换到查找侧栏。 */
  onOpenSearch?: () => void;
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
  reportPaginationDuration: (durationMs: number) => void,
  measurementReady: boolean,
) {
  const measureRef = useRef<HTMLDivElement>(null);
  const [measuredPages, setMeasuredPages] = useState<DocxPageContent[]>();

  useLayoutEffect(() => {
    setMeasuredPages(undefined);
    if (!measurementReady) return;
    const articles = Array.from(
      measureRef.current?.querySelectorAll<HTMLElement>(
        '.office-file-docx-page-frame__article',
      ) ?? [],
    );
    if (articles.length !== sourcePages.length) return;
    const startedAt = performance.now();

    const measured: DocxPageContent[] = [];
    sourcePages.forEach((sourcePage, sourcePageIndex) => {
      const elements = Array.from(articles[sourcePageIndex].children).filter(
        (element) =>
          !(element as HTMLElement).classList.contains(
            'office-file-word-notes',
          ),
      ) as HTMLElement[];
      if (elements.length !== sourcePage.blocks.length) return;
      const measuredFootnotes = measureDocxFootnotes(
        sourcePage.footnotes ?? [],
        articles[sourcePageIndex],
        sourcePage.page,
      );
      const claimedNoteIds = new Set<string>();
      const measurements: DocxMeasuredBlock[] = sourcePage.blocks.map(
        (block, blockIndex) => {
          const element = elements[blockIndex];
          const blockHeight = measureVisibleDocxBlockHeight(
            elements,
            blockIndex,
          );
          const alignedBlock =
            block.type === 'paragraph'
              ? alignDocxInlineObjectParagraphToGrid(
                  element,
                  block,
                  blockHeight,
                  sourcePage.page.gridLineHeight,
                )
              : { block, height: blockHeight };
          const blockFootnotes = Array.from(
            new Set(
              collectDocxNoteReferences([block])
                .filter((reference) => reference.noteKind === 'footnote')
                .map((reference) => reference.noteId),
            ),
          ).flatMap((noteId) => {
            if (claimedNoteIds.has(noteId)) return [];
            claimedNoteIds.add(noteId);
            const measured = measuredFootnotes.get(noteId);
            return measured ? [measured] : [];
          });
          return {
            block: alignedBlock.block,
            height: alignedBlock.height,
            paragraphBoxHeight:
              block.type === 'paragraph'
                ? element.getBoundingClientRect().height
                : undefined,
            leadingSpacing: Number.parseFloat(
              window.getComputedStyle(element).marginTop || '0',
            ),
            originalTableRowCount:
              block.type === 'table' ? block.rows.length : undefined,
            footnoteReserveHeight: blockFootnotes.reduce(
              (height, note) => height + (note.fragments[0]?.height ?? 0),
              0,
            ),
            measuredFootnotes: blockFootnotes,
            ...measureDocxParagraphLines(
              element,
              alignedBlock.block,
              alignedBlock.height,
            ),
            ...measureDocxTableRows(element, block),
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
  }, [measurementReady, reportPaginationDuration, sourcePages]);

  return { measureRef, pages: measuredPages ?? sourcePages };
}

// DocxViewer 负责 DOCX 页面内容的缩放渲染和滚动布局。
/** 协调 DOCX 页面、大纲和按需内容加载。 */
function DocxViewerComponent({
  preview,
  zoom,
  showOutline,
  wordRevisionMode,
  onCloseOutline,
  onOpenSearch,
}: DocxViewerProps) {
  const messages = useOfficeFileViewerMessages();
  const embeddedFontsReady = useOfficeEmbeddedFontsReady();
  const document =
    preview.mode === 'materialized' ? preview.model.document : undefined;
  const source = preview.mode === 'source' ? preview.source : undefined;
  const summary = preview.mode === 'source' ? preview.summary : undefined;
  const documentSessionId = preview.sessionId;
  const review = source ? summary?.review : document?.review;
  const notes = source ? summary?.notes : document?.notes;
  const noteByKey = useMemo(
    () =>
      new Map(
        [...(notes?.footnotes ?? []), ...(notes?.endnotes ?? [])].map(
          (note) => [`${note.kind}:${note.id}`, note] as const,
        ),
      ),
    [notes],
  );
  const annotationSource = useMemo(
    () =>
      review
        ? createMemoryOfficeAnnotationSource({
            annotations: review.annotations,
            revisions: review.revisions,
            revisionCount: review.revisionCount,
            noteCount: review.noteCount,
            supportsRevisionModes: review.supportsRevisionModes,
          })
        : undefined,
    [review],
  );
  useOfficeAnnotationSourceRegistration(annotationSource);
  const shouldRenderOutline = useWordOutlinePresence(
    showOutline,
    documentSessionId,
  );
  const scrollContainerRef = useRef<HTMLDivElement>(null);
  const materializedSourcePages = useMemo<DocxPageContent[]>(() => {
    if (!document || source) return [];
    const pages = document.pages?.length
      ? document.pages
      : [
          {
            id: 'docx-page-1',
            page: document.page,
            blocks: document.blocks,
          },
        ];
    return pages.map((page) => projectDocxPageContent(page, wordRevisionMode));
  }, [document, source, wordRevisionMode]);
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
    reportPaginationDuration,
    embeddedFontsReady,
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
        searchBlocks: document
          ? collectDocxSearchBlocks(
              materializedPages.flatMap((page) => page.blocks),
            )
          : undefined,
      }),
    [document, materializedPages],
  );
  useEffect(
    () => () => void materializedPageSource.dispose(),
    [materializedPageSource],
  );
  const pageSource = source ?? materializedPageSource;
  useOfficeSearchProviderRegistration(pageSource.searchProvider);
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
  useWordTargetNavigation({
    bookmarks: source ? summary?.bookmarks : document?.bookmarks,
    scrollContainerRef,
    pageMode: profile.pageMode,
    pageSource,
    blockPageIndex,
    pageNavigationControllerRef,
    documentSessionId,
  });
  useWordSearchNavigation({
    scrollContainerRef,
    pageMode: profile.pageMode,
    pageSource,
    blockPageIndex,
    pageNavigationControllerRef,
    documentSessionId,
  });
  useWordAnnotationNavigation({
    scrollContainerRef,
    pageMode: profile.pageMode,
    pageSource,
    blockPageIndex,
    pageNavigationControllerRef,
    documentSessionId,
  });
  const layoutKey = useMemo(
    () =>
      source
        ? `${documentSessionId}:${zoom}:${wordRevisionMode}:source`
        : `${zoom}:${wordRevisionMode}:${materializedPages
            .map((item) => item.id)
            .join('|')}`,
    [documentSessionId, materializedPages, source, wordRevisionMode, zoom],
  );

  const renderPageBlocks = useCallback(
    (pageItem: DocxPageContent, suppressFirstBlockSpacing = false) => {
      const contentWidth =
        pageItem.page.width -
        pageItem.page.marginLeft -
        pageItem.page.marginRight;
      return pageItem.blocks.map((block, blockIndex, blocks) => {
        const previousBlock = blocks[blockIndex - 1];
        const nextBlock = blocks[blockIndex + 1];
        const suppressSpacingBefore =
          (suppressFirstBlockSpacing && blockIndex === 0) ||
          shouldSuppressDocxContextualSpacing(block, previousBlock);
        return (
          <DocxBlockRenderer
            key={block.id}
            block={block}
            availableWidth={contentWidth}
            maximumWidth={pageItem.page.width}
            suppressSpacingBefore={suppressSpacingBefore}
            spacingBefore={resolveDocxSpacingBefore(
              block,
              previousBlock,
              suppressSpacingBefore,
            )}
            suppressSpacingAfter={shouldSuppressDocxContextualSpacing(
              block,
              nextBlock,
            )}
          />
        );
      });
    },
    [],
  );
  const renderPage = useCallback(
    (pageItem: DocxPageContent, pageIndex: number) => {
      const visiblePage = source
        ? projectDocxPageContent(pageItem, wordRevisionMode)
        : pageItem;
      const differentEvenOdd = Boolean(
        visiblePage.headers?.even !== undefined ||
          visiblePage.footerPageNumbers?.even !== undefined,
      );
      const headerBlocks = selectPageRegion<DocxPageContent['blocks']>(
        visiblePage.headers,
        pageIndex,
        visiblePage.differentFirstPage,
        differentEvenOdd,
      );
      const footerPageNumber = selectPageRegion<boolean>(
        visiblePage.footerPageNumbers,
        pageIndex,
        visiblePage.differentFirstPage,
        differentEvenOdd,
      );
      const displayedPageNumber =
        pageIndex + (visiblePage.differentFirstPage ? 0 : 1);
      const pageFootnotes =
        visiblePage.footnotes ??
        Array.from(
          new Set(
            collectDocxNoteReferences(visiblePage.blocks)
              .filter((reference) => reference.noteKind === 'footnote')
              .map((reference) => reference.noteId),
          ),
        ).flatMap((noteId) => {
          const note = noteByKey.get(`footnote:${noteId}`);
          return note ? [note] : [];
        });
      const footnotes = pageFootnotes.map((note) => ({
        ...note,
        blocks: note.blocks.map((block) =>
          projectDocxBlock(block, wordRevisionMode),
        ),
      }));
      const endnotes =
        pageIndex === pageSnapshot.pages.length - 1
          ? (notes?.endnotes ?? []).map((note) => ({
              ...note,
              blocks: note.blocks.map((block) =>
                projectDocxBlock(block, wordRevisionMode),
              ),
            }))
          : [];
      const showReviewOverlay = Boolean(
        review?.annotations.length ||
          (wordRevisionMode === 'markup' && review?.revisionCount),
      );
      return (
        <DocxPageFrame
          key={visiblePage.id}
          page={visiblePage.page}
          zoom={zoom}
          markupRailWidth={
            showReviewOverlay && wordRevisionMode === 'markup'
              ? WORD_MARKUP_BASE_RAIL_WIDTH
              : 0
          }
          header={
            headerBlocks?.length
              ? renderPageBlocks({ ...visiblePage, blocks: headerBlocks }, true)
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
          {renderPageBlocks(visiblePage, pageIndex > 0)}
          {showReviewOverlay ? <WordReviewOverlay /> : null}
          <WordNoteBlock notes={footnotes} page={visiblePage.page} />
          <WordNoteBlock notes={endnotes} page={visiblePage.page} endnotes />
        </DocxPageFrame>
      );
    },
    [
      noteByKey,
      notes?.endnotes,
      pageSnapshot.pages.length,
      renderPageBlocks,
      review,
      source,
      wordRevisionMode,
      zoom,
    ],
  );
  const measurementBatch = source?.getMeasurementBatch();
  const renderMeasurementBlock = useCallback(
    (
      block: DocxPageContent['blocks'][number],
      suppressSpacingBefore: boolean,
      suppressSpacingAfter: boolean,
      spacingBefore: number,
    ) => {
      const page = source?.getMeasurementBatch()?.sourcePage;
      const availableWidth = page
        ? page.page.width - page.page.marginLeft - page.page.marginRight
        : 0;
      const visibleBlock = projectDocxBlock(block, wordRevisionMode);
      return (
        <DocxBlockRenderer
          key={visibleBlock.id}
          block={visibleBlock}
          availableWidth={availableWidth}
          maximumWidth={page?.page.width}
          suppressSpacingBefore={suppressSpacingBefore}
          suppressSpacingAfter={suppressSpacingAfter}
          spacingBefore={spacingBefore}
        />
      );
    },
    [source, wordRevisionMode],
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
    return <OfficePreviewEmpty kind="docx" />;
  }
  if (!embeddedFontsReady) {
    return (
      <div className="office-file-docx-viewer office-file-docx-viewer--font-loading">
        <OfficeSpinner size="large" label={messages.loading.parsing} />
      </div>
    );
  }

  return (
    <WordRevisionModeContext.Provider value={wordRevisionMode}>
      <DocxCharacterSpacingContext.Provider value={compressPunctuation}>
        <div
          className="office-file-docx-viewer"
          data-word-source-mode={source ? 'progressive' : 'materialized'}
          data-character-spacing-control={characterSpacingControl}
        >
          <OfficeImagePreviewContext.Provider value={null}>
            {/* 隐藏测量副本不应注册图片交互，避免重复资源与焦点节点。 */}
            {source ? (
              <Suspense fallback={null}>
                <LazyDocxMeasureHost
                  batch={measurementBatch}
                  renderBlock={renderMeasurementBlock}
                  onMeasured={handleMeasured}
                  onError={handleMeasurementError}
                />
              </Suspense>
            ) : (
              <div
                ref={measureRef}
                className="office-file-docx-viewer__measure"
                aria-hidden="true"
              >
                {materializedSourcePages.map((pageItem) => (
                  <DocxPageFrame
                    key={pageItem.id}
                    page={pageItem.page}
                    zoom={100}
                  >
                    {renderPageBlocks(pageItem)}
                    <WordNoteBlock
                      notes={pageItem.footnotes ?? []}
                      page={pageItem.page}
                    />
                  </DocxPageFrame>
                ))}
              </div>
            )}
          </OfficeImagePreviewContext.Provider>
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
              onOpenSearch={onOpenSearch}
            />
            <OfficeWatermarkSurface>
              <div
                ref={scrollContainerRef}
                className="office-file-docx-viewer__scroller"
                data-office-fit-viewport="true"
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
            </OfficeWatermarkSurface>
          </div>
        </div>
      </DocxCharacterSpacingContext.Provider>
    </WordRevisionModeContext.Provider>
  );
}

export const DocxViewer = memo(DocxViewerComponent);
