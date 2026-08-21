import React, {
  memo,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import {
  createMemoryOfficeAnnotationSource,
  type WordRevisionMode,
} from '../../services/annotations';
import {
  pageDrawingImagesFromBlock,
  paginateDocBlocks,
  type PaginatedDocPage,
} from '../../services/doc/docPagination';
import { projectDocBlocksRevisionMode } from '../../services/doc/projectDocRevisionMode';
import { docBookmarkMarkerIdsFromBlock } from '../../services/doc/readDocBookmarks';
import type { DocDocument, DocNote } from '../../services/doc/types';
import type { OfficeFileViewerPreviewState } from '../../services/parsing/internalTypes';
import { collectWordPerformanceStats } from '../../services/word/collectWordPerformanceStats';
import { createMaterializedWordPageSource } from '../../services/word/createMaterializedWordPageSource';
import { createMemoryWordOutlineProvider } from '../../services/word/createMemoryWordOutlineProvider';
import { collectDocSearchBlocks } from '../../services/word/WordSearchProvider';
import { useOfficeAnnotationSourceRegistration } from '../../shared/annotations';
import { useExternalStoreSnapshot } from '../../shared/react/useExternalStoreSnapshot';
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
import { WordReviewOverlay } from '../word-review/WordReviewOverlay';
import { WordRevisionModeContext } from '../word-review/WordRevisionText';
import { useWordSearchNavigation } from '../word-search/useWordSearchNavigation';
import { DocContentRenderer } from './DocContentRenderer';
import { DocImageGallery } from './DocImageGallery';
import { collectDocNoteReferences, DocNoteBlock } from './DocNoteBlock';
import { DocPageFrame } from './DocPageFrame';
import './index.less';

/** DOC/WPS Viewer 可以消费的物化或按需预览。 */
type DocPreview = Extract<OfficeFileViewerPreviewState, { previewKind: 'doc' }>;

/** DOC预览器组件属性。 */
type DocViewerProps = {
  /** 当前 DOC/WPS 的物化或按需预览。 */
  preview: DocPreview;
  /** 当前预览缩放比例，100 表示原始大小。 */
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
/** 浏览器真实排版与静态估算不一致时记录的 DOC 块高度。 */
type DocLayoutCalibration = {
  /** 当前校准结果所属的解析会话。 */
  sessionId: string;
  /** 按原始块 ID 保存的真实外框高度。 */
  blockHeights: ReadonlyMap<string, number>;
};

/** 收集普通 DOC 模型中已锚定的图片，避免尾页 Gallery 重复显示。 */
function collectAnchoredImageIds(document: DocDocument | undefined) {
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

/** 从分页结果定位正文页码起点；目录页使用自身编号但不显示正文页脚。 */
function resolveFooterPageNumberStartIndex(pages: PaginatedDocPage[]) {
  const lastTocPageIndex = pages.reduce(
    (lastIndex, currentPage, pageIndex) =>
      currentPage.blocks.some(
        (block) => block.type === 'paragraph' && block.isTableOfContents,
      )
        ? pageIndex
        : lastIndex,
    -1,
  );
  // 无目录时保持首页封面不编号的既有规则，避免普通 DOC 回归时新增页脚。
  return lastTocPageIndex >= 0 ? lastTocPageIndex + 1 : 1;
}

/** 渲染普通 DOC 模型或大文件渐进 PageSource。 */
function DocViewerComponent({
  preview,
  zoom,
  showOutline,
  wordRevisionMode,
  onCloseOutline,
  onOpenSearch,
}: DocViewerProps) {
  const rawDocument =
    preview.mode === 'materialized' ? preview.model.document : undefined;
  const document = useMemo(
    () =>
      rawDocument
        ? {
            ...rawDocument,
            blocks: projectDocBlocksRevisionMode(
              rawDocument.blocks,
              wordRevisionMode,
            ),
          }
        : undefined,
    [rawDocument, wordRevisionMode],
  );
  const source = preview.mode === 'source' ? preview.source : undefined;
  const summary = preview.mode === 'source' ? preview.summary : undefined;
  const documentSessionId = preview.sessionId;
  const shouldRenderOutline = useWordOutlinePresence(
    showOutline,
    documentSessionId,
  );
  const scrollContainerRef = useRef<HTMLDivElement>(null);
  const [layoutCalibration, setLayoutCalibration] =
    useState<DocLayoutCalibration>(() => ({
      sessionId: '',
      blockHeights: new Map(),
    }));
  const calibratedBlockHeights =
    layoutCalibration.sessionId === documentSessionId
      ? layoutCalibration.blockHeights
      : undefined;
  const documentMetadata = summary ?? document;
  const notes = documentMetadata?.notes;
  const review = documentMetadata?.review;
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
  const noteByKey = useMemo<Map<string, DocNote>>(
    () =>
      new Map(
        [...(notes?.footnotes ?? []), ...(notes?.endnotes ?? [])].map(
          (note) => [`${note.noteKind}:${note.noteId}`, note] as const,
        ),
      ),
    [notes],
  );
  const page = documentMetadata?.page;
  const contentWidth = page
    ? page.width - page.marginLeft - page.marginRight
    : 0;
  const pagination = useMemo(() => {
    const startedAt = performance.now();
    const nextPages =
      document && !source && page
        ? paginateDocBlocks(
            document.blocks,
            page,
            contentWidth,
            calibratedBlockHeights,
          )
        : [];
    return {
      pages: nextPages,
      durationMs: performance.now() - startedAt,
    };
  }, [calibratedBlockHeights, contentWidth, document, page, source]);
  const materializedPages = pagination.pages;
  const footerPageNumberStartIndex = useMemo(
    () => resolveFooterPageNumberStartIndex(materializedPages),
    [materializedPages],
  );
  const materializedPageSource = useMemo(
    () =>
      createMaterializedWordPageSource(materializedPages, {
        getId: (docPage) => docPage.id,
        getEstimatedContentHeight: () => page?.minHeight ?? 0,
        getSourceBlockIds: (docPage) =>
          docPage.blocks.flatMap((block) => [
            block.id,
            ...(block.sourceBlockId ? [block.sourceBlockId] : []),
            ...docBookmarkMarkerIdsFromBlock(block),
          ]),
        searchBlocks: document
          ? collectDocSearchBlocks(document.blocks)
          : undefined,
      }),
    [document, materializedPages, page?.minHeight],
  );
  useEffect(
    () => () => void materializedPageSource.dispose(),
    [materializedPageSource],
  );
  const pageSource = source?.pages ?? materializedPageSource;
  useOfficeSearchProviderRegistration(pageSource.searchProvider);
  const pageSnapshot = useExternalStoreSnapshot(pageSource);
  const blockPageIndex = useMemo(() => {
    const index = new WordBlockPageIndex();
    pageSnapshot.pages.forEach((meta) => index.replacePage(meta));
    return index;
  }, [pageSnapshot]);
  const pageNavigationControllerRef = useRef<WordPageNavigationController>();
  const outlineItems = useMemo(
    () =>
      shouldRenderOutline
        ? source?.getOutlineItems() ?? document?.outline ?? []
        : [],
    [document?.outline, pageSnapshot.revision, shouldRenderOutline, source],
  );
  const memoryOutlineProvider = useMemo(
    () => createMemoryWordOutlineProvider(outlineItems),
    [outlineItems],
  );
  const outlineProvider = source?.outline ?? memoryOutlineProvider;
  const performanceStats = useMemo(
    () =>
      document
        ? collectWordPerformanceStats(document, {
            estimatedPageCount: materializedPages.length,
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
    [document, materializedPages.length],
  );
  const { profile: materializedProfile, reportPaginationDuration } =
    useWordPerformanceProfile(documentSessionId, performanceStats);
  const profile = source?.getPerformanceProfile() ?? materializedProfile;
  useLayoutEffect(() => {
    if (source || profile.pageMode === 'windowed') return;
    const container = scrollContainerRef.current;
    if (!container) return;

    const scale = Math.max(zoom / 100, 0.01);
    const nextBlockHeights = new Map(calibratedBlockHeights ?? []);
    let changed = false;
    container
      .querySelectorAll<HTMLElement>('[data-office-doc-pagination-id]')
      .forEach((element) => {
        // 图片有独立的尺寸估算与加载时机，只校准浏览器字体实际换行造成的文字高度差。
        if (element.querySelector('img')) return;
        if (element.dataset.officeDocPaginationFragment === 'true') return;
        const blockId = element.dataset.officeDocPaginationId;
        const estimatedHeight = Number(
          element.dataset.officeDocEstimatedHeight,
        );
        if (!blockId || !Number.isFinite(estimatedHeight)) return;
        const style = window.getComputedStyle(element);
        const actualHeight =
          element.getBoundingClientRect().height / scale +
          (Number.parseFloat(style.marginTop) || 0) +
          (Number.parseFloat(style.marginBottom) || 0);
        if (actualHeight <= estimatedHeight + 0.75) return;
        const calibratedHeight = Math.ceil(actualHeight * 10) / 10;
        if (calibratedHeight <= (nextBlockHeights.get(blockId) ?? 0) + 0.5) {
          return;
        }
        nextBlockHeights.set(blockId, calibratedHeight);
        changed = true;
      });
    if (changed) {
      setLayoutCalibration({
        sessionId: documentSessionId,
        blockHeights: nextBlockHeights,
      });
    }
  }, [
    calibratedBlockHeights,
    documentSessionId,
    materializedPages,
    profile.pageMode,
    source,
    zoom,
  ]);
  useWordTargetNavigation({
    bookmarks: documentMetadata?.bookmarks,
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
  useEffect(() => {
    if (!source) reportPaginationDuration(pagination.durationMs);
  }, [pagination.durationMs, reportPaginationDuration, source]);
  const anchoredImageIds = useMemo(
    () => collectAnchoredImageIds(document),
    [document],
  );
  const unanchoredImages = useMemo(
    () =>
      source
        ? []
        : document?.images.filter((image) => !anchoredImageIds.has(image.id)) ??
          [],
    [anchoredImageIds, document, source],
  );
  const layoutKey = useMemo(
    () =>
      `${zoom}:${wordRevisionMode}:${pageSnapshot.revision}:${pageSnapshot.pages
        .map((item) => item.id)
        .join('|')}`,
    [pageSnapshot, wordRevisionMode, zoom],
  );

  if (!page || !documentMetadata || !pageSnapshot.pages.length) {
    return <OfficePreviewEmpty kind="doc" />;
  }

  const renderPage = (docPage: PaginatedDocPage, pageIndex: number) => {
    const visiblePage = source
      ? {
          ...docPage,
          blocks: projectDocBlocksRevisionMode(
            docPage.blocks,
            wordRevisionMode,
          ),
        }
      : docPage;
    const seenFootnotes = new Set<string>();
    const footnotes = collectDocNoteReferences(visiblePage.blocks).flatMap(
      (reference) => {
        if (reference.noteKind !== 'footnote') return [];
        const key = `footnote:${reference.noteId}`;
        if (seenFootnotes.has(key)) return [];
        seenFootnotes.add(key);
        const note = noteByKey.get(key);
        return note ? [note] : [];
      },
    );
    const lastPageIndex =
      (pageSnapshot.pageCount ?? pageSnapshot.pages.length) - 1;
    const endnotes = pageIndex === lastPageIndex ? notes?.endnotes ?? [] : [];
    const showReviewOverlay = Boolean(
      review?.annotations.length ||
        (wordRevisionMode === 'markup' && review?.revisionCount),
    );
    return (
      <DocPageFrame
        key={visiblePage.id}
        page={page}
        zoom={zoom}
        markupRailWidth={
          showReviewOverlay && wordRevisionMode === 'markup'
            ? WORD_MARKUP_BASE_RAIL_WIDTH
            : 0
        }
        headerImage={documentMetadata.headerImage}
        pageDrawings={visiblePage.blocks.flatMap(pageDrawingImagesFromBlock)}
        footerText={
          documentMetadata.footerPageNumbers &&
          pageIndex >= footerPageNumberStartIndex
            ? `${pageIndex - footerPageNumberStartIndex + 1}`
            : undefined
        }
      >
        <DocContentRenderer
          blocks={visiblePage.blocks}
          contentWidth={contentWidth}
        />
        {showReviewOverlay ? <WordReviewOverlay /> : null}
        <DocNoteBlock notes={footnotes} page={page} />
        <DocNoteBlock notes={endnotes} page={page} endnotes />
        {pageIndex === lastPageIndex ? (
          <DocImageGallery images={unanchoredImages} />
        ) : null}
      </DocPageFrame>
    );
  };
  return (
    <WordRevisionModeContext.Provider value={wordRevisionMode}>
      <div
        className="office-file-doc-viewer"
        data-word-source-mode={source ? 'progressive' : 'materialized'}
      >
        <div className="office-file-doc-viewer__body">
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
          <div
            ref={scrollContainerRef}
            className="office-file-doc-viewer__scroller"
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
        </div>
      </div>
    </WordRevisionModeContext.Provider>
  );
}

export const DocViewer = memo(DocViewerComponent);
