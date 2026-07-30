import React, {
  memo,
  useEffect,
  useMemo,
  useRef,
  useSyncExternalStore,
} from 'react';
import type {
  DocWordPageSource,
  DocWordPreviewSummary,
} from '../../services/doc/DocWordPageSource';
import type { DocDocument } from '../../services/doc/types';
import { collectWordPerformanceStats } from '../../services/word/collectWordPerformanceStats';
import { createMaterializedWordPageSource } from '../../services/word/createMaterializedWordPageSource';
import { createMemoryWordOutlineProvider } from '../../services/word/createMemoryWordOutlineProvider';
import { OfficeEmpty } from '../../shell/Empty';
import { WordOutlineSidebar } from '../word-outline/WordOutlineSidebar';
import type { WordPageNavigationController } from '../word-pages/types';
import { VirtualWordPageList } from '../word-pages/VirtualWordPageList';
import { WordBlockPageIndex } from '../word-pages/WordBlockPageIndex';
import { useWordPerformanceProfile } from '../word-performance/useWordPerformanceProfile';
import { DocContentRenderer } from './DocContentRenderer';
import { DocImageGallery } from './DocImageGallery';
import { DocPageFrame } from './DocPageFrame';
import { paginateDocBlocks, type PaginatedDocPage } from './docRenderUtils';
import './index.less';

/** 定义 DocViewer 组件可接收的属性。 */
type DocViewerProps = {
  document?: DocDocument;
  source?: DocWordPageSource;
  summary?: DocWordPreviewSummary;
  zoom: number;
  documentSessionId: string;
};

/** 收集普通 DOC 模型中已锚定的图片，避免尾页 Gallery 重复显示。 */
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

/** 渲染普通 DOC 模型或大文件渐进 PageSource。 */
function DocViewerComponent({
  document,
  source,
  summary,
  zoom,
  documentSessionId,
}: DocViewerProps) {
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
  const materializedPageSource = useMemo(
    () =>
      createMaterializedWordPageSource(materializedPages, {
        getId: (docPage) => docPage.id,
        getEstimatedContentHeight: () => page?.minHeight ?? 0,
        getSourceBlockIds: (docPage) =>
          docPage.blocks.flatMap((block) => [
            block.id,
            ...(block.sourceBlockId ? [block.sourceBlockId] : []),
          ]),
      }),
    [materializedPages, page?.minHeight],
  );
  useEffect(
    () => () => void materializedPageSource.dispose(),
    [materializedPageSource],
  );
  const pageSource = source?.pages ?? materializedPageSource;
  const pageSnapshot = useSyncExternalStore(
    pageSource.subscribe,
    pageSource.getSnapshot,
    pageSource.getSnapshot,
  );
  const blockPageIndex = useMemo(() => {
    const index = new WordBlockPageIndex();
    pageSnapshot.pages.forEach((meta) => index.replacePage(meta));
    return index;
  }, [pageSnapshot]);
  const pageNavigationControllerRef = useRef<WordPageNavigationController>();
  const outlineItems = useMemo(
    () => source?.getOutlineItems() ?? document?.outline ?? [],
    [document?.outline, pageSnapshot.revision, source],
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
    return <OfficeEmpty kind="doc" />;
  }

  const renderPage = (docPage: PaginatedDocPage, pageIndex: number) => (
    <DocPageFrame
      key={docPage.id}
      page={page}
      zoom={zoom}
      headerImage={documentMetadata.headerImage}
      footerText={
        documentMetadata.footerPageNumbers && pageIndex > 0
          ? `- ${pageIndex} -`
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
      {documentMetadata.warnings.length ? (
        <div className="office-file-doc-viewer__notice" role="alert">
          {documentMetadata.warnings.join(' ')}
        </div>
      ) : null}
      <div className="office-file-doc-viewer__body">
        <WordOutlineSidebar
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
