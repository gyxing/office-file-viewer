import React, { memo, useEffect, useMemo, useRef } from 'react';
import {
  pageDrawingImagesFromBlock,
  paginateDocBlocks,
  type PaginatedDocPage,
} from '../../services/doc/docPagination';
import { docBookmarkMarkerIdsFromBlock } from '../../services/doc/readDocBookmarks';
import type { DocDocument } from '../../services/doc/types';
import type { OfficeFileViewerPreviewState } from '../../services/parsing/internalTypes';
import { collectWordPerformanceStats } from '../../services/word/collectWordPerformanceStats';
import { createMaterializedWordPageSource } from '../../services/word/createMaterializedWordPageSource';
import { createMemoryWordOutlineProvider } from '../../services/word/createMemoryWordOutlineProvider';
import { useExternalStoreSnapshot } from '../../shared/react/useExternalStoreSnapshot';
import { OfficePreviewEmpty } from '../common/OfficePreviewEmpty';
import { useWordTargetNavigation } from '../word-hyperlink/useWordTargetNavigation';
import { useWordOutlinePresence } from '../word-outline/useWordOutlinePresence';
import { WordOutlineSidebar } from '../word-outline/WordOutlineSidebar';
import type { WordPageNavigationController } from '../word-pages/types';
import { VirtualWordPageList } from '../word-pages/VirtualWordPageList';
import { WordBlockPageIndex } from '../word-pages/WordBlockPageIndex';
import { useWordPerformanceProfile } from '../word-performance/useWordPerformanceProfile';
import { DocContentRenderer } from './DocContentRenderer';
import { DocImageGallery } from './DocImageGallery';
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
  /** 关闭文档大纲。 */
  onCloseOutline: () => void;
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
  onCloseOutline,
}: DocViewerProps) {
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
  const documentMetadata = summary ?? document;
  const page = documentMetadata?.page;
  const contentWidth = page
    ? page.width - page.marginLeft - page.marginRight
    : 0;
  const pagination = useMemo(() => {
    const startedAt = performance.now();
    const nextPages =
      document && !source && page
        ? paginateDocBlocks(document.blocks, page, contentWidth)
        : [];
    return {
      pages: nextPages,
      durationMs: performance.now() - startedAt,
    };
  }, [contentWidth, document, page, source]);
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
      }),
    [materializedPages, page?.minHeight],
  );
  useEffect(
    () => () => void materializedPageSource.dispose(),
    [materializedPageSource],
  );
  const pageSource = source?.pages ?? materializedPageSource;
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
  useWordTargetNavigation({
    bookmarks: documentMetadata?.bookmarks,
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
      `${zoom}:${pageSnapshot.revision}:${pageSnapshot.pages
        .map((item) => item.id)
        .join('|')}`,
    [pageSnapshot, zoom],
  );

  if (!page || !documentMetadata || !pageSnapshot.pages.length) {
    return <OfficePreviewEmpty kind="doc" />;
  }

  const renderPage = (docPage: PaginatedDocPage, pageIndex: number) => (
    <DocPageFrame
      key={docPage.id}
      page={page}
      zoom={zoom}
      headerImage={documentMetadata.headerImage}
      pageDrawings={docPage.blocks.flatMap(pageDrawingImagesFromBlock)}
      footerText={
        documentMetadata.footerPageNumbers &&
        pageIndex >= footerPageNumberStartIndex
          ? `${pageIndex - footerPageNumberStartIndex + 1}`
          : undefined
      }
    >
      <DocContentRenderer blocks={docPage.blocks} contentWidth={contentWidth} />
      {pageIndex ===
      (pageSnapshot.pageCount ?? pageSnapshot.pages.length) - 1 ? (
        <DocImageGallery images={unanchoredImages} />
      ) : null}
    </DocPageFrame>
  );

  return (
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
        />
        <div
          ref={scrollContainerRef}
          className="office-file-doc-viewer__scroller"
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
  );
}

export const DocViewer = memo(DocViewerComponent);
