// XlsxViewer 负责 XLSX 工作簿预览的整体布局，包括工作表选择和当前工作表内容区。
import React, {
  memo,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import { createMemoryOfficeAnnotationSource } from '../../services/annotations';
import type { OfficeFileViewerPreviewState } from '../../services/parsing/internalTypes';
import type { SpreadsheetSource } from '../../services/spreadsheet/SpreadsheetSource';
import { getSpreadsheetSource } from '../../services/spreadsheet/spreadsheetSourceRegistry';
import type { SpreadsheetViewMode } from '../../services/spreadsheet/viewMode';
import { useOfficeAnnotationSourceRegistration } from '../../shared/annotations';
import { OfficeWatermarkSurface } from '../../shared/watermark';
import { OfficePreviewEmpty } from '../common/OfficePreviewEmpty';
import { useOfficeSearchProviderRegistration } from '../search/OfficeSearchContext';
import './index.less';
import type { SpreadsheetNavigationController } from './spreadsheetNavigation';
import { SpreadsheetSheetState } from './SpreadsheetSheetState';
import { useSpreadsheetAnnotationNavigation } from './useSpreadsheetAnnotationNavigation';
import { useSpreadsheetHyperlinkNavigation } from './useSpreadsheetHyperlinkNavigation';
import { useSpreadsheetSearchNavigation } from './useSpreadsheetSearchNavigation';
import { useSpreadsheetSource } from './useSpreadsheetSource';
import { VirtualSpreadsheetGrid } from './VirtualSpreadsheetGrid';
import { XlsxChartSheet } from './XlsxChartSheet';
import { XlsxSheetGrid } from './XlsxSheetGrid';
import { XlsxSheetTabs } from './XlsxSheetTabs';

/** 尚未计算阅读行高时复用的空集合。 */
const EMPTY_READING_ROW_HEIGHTS: ReadonlyMap<number, number> = new Map();

/** 电子表格 Viewer 可以消费的物化或按需预览。 */
type SpreadsheetPreview = Extract<
  OfficeFileViewerPreviewState,
  { previewKind: 'xls' | 'xlsx' }
>;

/** Excel预览器组件属性。 */
type XlsxViewerProps = {
  /** 当前电子表格的物化或按需预览。 */
  preview: SpreadsheetPreview;
  /** 当前激活工作表的稳定标识。 */
  activeSheetId?: string;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 当前电子表格采用的显示模式。 */
  viewMode: SpreadsheetViewMode;
  /** 在 SelectSheet 事件发生时调用的回调函数。 */
  onSelectSheet: (sheetId: string) => void;
};

/** 渲染Excel预览器。 */
function XlsxViewerComponent({
  preview,
  activeSheetId,
  zoom,
  viewMode,
  onSelectSheet,
}: XlsxViewerProps) {
  const resolvedSource = useMemo(
    () =>
      preview.mode === 'source'
        ? preview.source
        : getSpreadsheetSource(preview.model.workbook),
    [preview],
  );

  if (!resolvedSource) {
    return <OfficePreviewEmpty kind={preview.previewKind} />;
  }

  return (
    <XlsxSourceViewer
      source={resolvedSource}
      kind={preview.previewKind}
      activeSheetId={activeSheetId}
      zoom={zoom}
      viewMode={viewMode}
      onSelectSheet={onSelectSheet}
    />
  );
}

/** 渲染已经统一为 Source 的工作簿。 */
function XlsxSourceViewer({
  source,
  kind,
  activeSheetId,
  zoom,
  viewMode,
  onSelectSheet,
}: {
  source: SpreadsheetSource;
  kind: 'xlsx' | 'xls';
  activeSheetId?: string;
  zoom: number;
  viewMode: SpreadsheetViewMode;
  onSelectSheet: (sheetId: string) => void;
}) {
  useOfficeSearchProviderRegistration(source.searchProvider);
  const readingRowHeightCacheRef = useRef(
    new Map<string, ReadonlyMap<number, number>>(),
  );
  const sourceRef = useRef(source);
  const [readingLayoutRevision, setReadingLayoutRevision] = useState(0);
  if (sourceRef.current !== source) {
    // 数据源变化后旧 Sheet 标识可能复用，必须同步丢弃此前的阅读布局缓存。
    sourceRef.current = source;
    readingRowHeightCacheRef.current.clear();
  }
  const state = useSpreadsheetSource(source, activeSheetId);
  const descriptor = state.activeDescriptor;
  const [annotations, setAnnotations] = useState<
    Awaited<ReturnType<SpreadsheetSource['getAnnotations']>>
  >([]);
  useEffect(() => {
    const controller = new AbortController();
    setAnnotations([]);
    if (!descriptor || descriptor.kind !== 'worksheet') {
      return () => controller.abort();
    }
    void source.getAnnotations(descriptor.id, controller.signal).then(
      (items) => {
        if (!controller.signal.aborted) setAnnotations(items);
      },
      () => undefined,
    );
    return () => controller.abort();
  }, [descriptor, source]);
  const annotationSource = useMemo(
    () =>
      descriptor && annotations.length
        ? createMemoryOfficeAnnotationSource({
            annotations: annotations.map((annotation) => ({
              id: annotation.id,
              author: annotation.author,
              createdAt: annotation.createdAt,
              text: annotation.text,
              resolved: annotation.resolved,
              parentId: annotation.parentId,
              target: {
                kind: 'spreadsheet-cell',
                sheetId: descriptor.id,
                row: annotation.row,
                column: annotation.column,
              },
            })),
          })
        : undefined,
    [annotations, descriptor],
  );
  useOfficeAnnotationSourceRegistration(annotationSource);
  const navigationControllerRef = useRef<SpreadsheetNavigationController>();
  useSpreadsheetHyperlinkNavigation({
    snapshot: state.snapshot,
    activeSheetId: descriptor?.id,
    onSelectSheet,
    navigationControllerRef,
  });
  useSpreadsheetSearchNavigation({
    activeSheetId: descriptor?.id,
    onSelectSheet,
    navigationControllerRef,
  });
  useSpreadsheetAnnotationNavigation({
    activeSheetId: descriptor?.id,
    onSelectSheet,
    navigationControllerRef,
  });
  const readingRowHeights = useMemo(
    () =>
      descriptor
        ? readingRowHeightCacheRef.current.get(descriptor.id) ??
          EMPTY_READING_ROW_HEIGHTS
        : EMPTY_READING_ROW_HEIGHTS,
    [descriptor, readingLayoutRevision, source],
  );
  const handleReadingRowHeightsChange = useCallback(
    (sheetId: string, updates: ReadonlyMap<number, number>) => {
      if (!updates.size) return;
      const current =
        readingRowHeightCacheRef.current.get(sheetId) ??
        EMPTY_READING_ROW_HEIGHTS;
      let next: Map<number, number> | undefined;
      updates.forEach((height, rowIndex) => {
        if (height <= (current.get(rowIndex) ?? 0) + 0.5) return;
        next ??= new Map(current);
        next.set(rowIndex, height);
      });
      if (!next) return;
      readingRowHeightCacheRef.current.set(sheetId, next);
      setReadingLayoutRevision((revision) => revision + 1);
    },
    [],
  );

  if (!descriptor) return <OfficePreviewEmpty kind={kind} />;

  return (
    <div className="office-file-xlsx-viewer">
      <XlsxSheetTabs
        snapshot={state.snapshot}
        activeSheet={descriptor}
        onSelectSheet={onSelectSheet}
      />
      <OfficeWatermarkSurface>
        <SpreadsheetSheetState
          loading={state.loading}
          error={state.error}
          retry={state.retry}
        >
          {state.activeSheet?.kind === 'chart' ? (
            <XlsxChartSheet sheet={state.activeSheet} zoom={zoom} />
          ) : state.profile?.gridMode === 'table' && state.activeSheet ? (
            <XlsxSheetGrid
              sheet={state.activeSheet}
              zoom={zoom}
              viewMode={viewMode}
              navigationControllerRef={navigationControllerRef}
            />
          ) : state.profile ? (
            <VirtualSpreadsheetGrid
              source={source}
              sheetId={descriptor.id}
              layout={source.getSheetLayout(descriptor.id)}
              gridMode={state.profile.gridMode}
              zoom={zoom}
              viewMode={viewMode}
              readingRowHeights={readingRowHeights}
              onReadingRowHeightsChange={handleReadingRowHeightsChange}
              navigationControllerRef={navigationControllerRef}
            />
          ) : null}
        </SpreadsheetSheetState>
      </OfficeWatermarkSurface>
    </div>
  );
}

export const XlsxViewer = memo(XlsxViewerComponent);
