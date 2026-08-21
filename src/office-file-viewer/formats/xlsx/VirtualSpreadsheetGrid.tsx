// VirtualSpreadsheetGrid 按完整工作表坐标渲染当前二维窗口。
import type { CSSProperties, MutableRefObject } from 'react';
import React, {
  memo,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type { OfficeFontFamilyResolver } from '../../services/fonts/types';
import {
  useOfficeResourceUrl,
  type OfficeResourceSource,
} from '../../services/resource-store';
import type { SpreadsheetAxisIndex } from '../../services/spreadsheet/SpreadsheetAxisIndex';
import type { SpreadsheetPerformanceProfile } from '../../services/spreadsheet/spreadsheetPerformance';
import type {
  SpreadsheetSheetLayout,
  SpreadsheetSource,
} from '../../services/spreadsheet/SpreadsheetSource';
import type {
  SpreadsheetCell,
  SpreadsheetChart,
  SpreadsheetImage,
  SpreadsheetMerge,
  SpreadsheetRangeData,
} from '../../services/spreadsheet/types';
import type { SpreadsheetViewMode } from '../../services/spreadsheet/viewMode';
import { OfficeChartView } from '../../shared/chart/OfficeChartView';
import { useOfficeFontResolver } from '../../shared/fonts/OfficeFontProvider';
import { useOfficeHyperlink } from '../../shared/hyperlink';
import { OfficePreviewableImage } from '../../shared/image-preview';
import {
  buildXlsxCellStyle,
  getSpreadsheetColumnLabel,
  isHighlightedXlsxCell,
  XLSX_COLUMN_HEADER_HEIGHT,
  XLSX_ROW_HEADER_WIDTH,
} from './sheetRenderUtils';
import {
  buildSpreadsheetRowContentBounds,
  isSpreadsheetCellOccupied,
  type SpreadsheetCellContentBounds,
} from './spreadsheetCellOverflow';
import { SpreadsheetCellRenderer } from './SpreadsheetCellRenderer';
import {
  SpreadsheetFrozenPanes,
  type SpreadsheetFrozenCellRenderArgs,
} from './SpreadsheetFrozenPanes';
import { SpreadsheetGridPlaceholder } from './SpreadsheetGridPlaceholder';
import type { SpreadsheetNavigationController } from './spreadsheetNavigation';
import {
  buildSpreadsheetReadingRowHeightUpdates,
  isSpreadsheetShrinkToFitCell,
  remapSpreadsheetVerticalRange,
} from './spreadsheetReadingLayout';
import { useSpreadsheetGridWindow } from './useSpreadsheetGridWindow';

/** SSR 环境延后布局副作用，浏览器中则在绘制前提交阅读行高。 */
const useIsomorphicLayoutEffect =
  typeof window === 'undefined' ? useEffect : useLayoutEffect;
/** 网格样式保留的未缩放内边距；滚动后标题层需抵消该区域。 */
const XLSX_GRID_VIEWPORT_INSET = 16;

/** 电子表格虚拟网格组件属性。 */
type VirtualSpreadsheetGridProps = {
  /** 当前预览使用的按需加载数据源。 */
  source: SpreadsheetSource;
  /** 工作表的稳定标识。 */
  sheetId: string;
  /** 当前内容使用的布局信息。 */
  layout: SpreadsheetSheetLayout;
  /** 当前工作表使用的网格渲染模式。 */
  gridMode: SpreadsheetPerformanceProfile['gridMode'];
  /** 当前预览缩放比例，100 表示原始大小。 */
  zoom: number;
  /** 当前电子表格采用的显示模式。 */
  viewMode: SpreadsheetViewMode;
  /** 当前工作表已经计算出的稀疏阅读行高。 */
  readingRowHeights: ReadonlyMap<number, number>;
  /** 合并当前窗口新计算出的阅读行高。 */
  onReadingRowHeightsChange: (
    sheetId: string,
    updates: ReadonlyMap<number, number>,
  ) => void;
  /** 供工作簿内部链接驱动当前虚拟网格定位。 */
  navigationControllerRef: MutableRefObject<
    SpreadsheetNavigationController | undefined
  >;
};

/** 电子表格已经加载的行列范围。 */
type LoadedSpreadsheetRange = {
  /** 当前预览使用的按需加载数据源。 */
  source: SpreadsheetSource;
  /** 工作表的稳定标识。 */
  sheetId: string;
  /** 已加载范围内的行列与单元格数据。 */
  data: SpreadsheetRangeData;
};

function positionKey(row: number, column: number) {
  return `${row}:${column}`;
}

/** 在虚拟画布中按需加载单个图片资源。 */
function VirtualSpreadsheetImage({
  image,
  sourceRowAxis,
  rowAxis,
}: {
  image: SpreadsheetImage;
  sourceRowAxis: SpreadsheetAxisIndex;
  rowAxis: SpreadsheetAxisIndex;
}) {
  const messages = useOfficeFileViewerMessages();
  const source = useMemo<OfficeResourceSource>(
    () =>
      typeof image.src === 'string'
        ? { kind: 'url', url: image.src }
        : image.src,
    [image.src],
  );
  const resource = useOfficeResourceUrl(source);
  const hyperlinkProps = useOfficeHyperlink<HTMLImageElement>({
    hyperlink: image.hyperlink,
    source: { type: 'image', id: image.id },
  });
  const verticalRange = useMemo(
    () =>
      remapSpreadsheetVerticalRange(
        image.y,
        image.height,
        sourceRowAxis,
        rowAxis,
      ),
    [image.height, image.y, rowAxis, sourceRowAxis],
  );
  return (
    <OfficePreviewableImage
      {...hyperlinkProps}
      previewId={image.id}
      previewName={image.name}
      previewSource={source}
      className="office-file-xlsx-sheet-grid__floating-image"
      src={resource.url}
      alt={image.alt ?? ''}
      title={image.name}
      aria-busy={resource.loading || undefined}
      aria-label={
        resource.error
          ? messages.spreadsheet.imageLoadFailed(image.alt)
          : undefined
      }
      data-load-error={resource.error ? 'true' : undefined}
      style={{
        left: XLSX_ROW_HEADER_WIDTH + image.x,
        top: XLSX_COLUMN_HEADER_HEIGHT + verticalRange.y,
        width: image.width,
        height: verticalRange.height,
      }}
    />
  );
}

/** 在虚拟画布中按全局锚点渲染单个图表。 */
function VirtualSpreadsheetChart({
  chart,
  zoom,
  sourceRowAxis,
  rowAxis,
}: {
  chart: SpreadsheetChart;
  zoom: number;
  sourceRowAxis: SpreadsheetAxisIndex;
  rowAxis: SpreadsheetAxisIndex;
}) {
  const hyperlinkProps = useOfficeHyperlink<HTMLDivElement>({
    hyperlink: chart.hyperlink,
    source: { type: 'shape', id: chart.id },
  });
  const verticalRange = useMemo(
    () =>
      remapSpreadsheetVerticalRange(
        chart.y,
        chart.height,
        sourceRowAxis,
        rowAxis,
      ),
    [chart.height, chart.y, rowAxis, sourceRowAxis],
  );
  return (
    <div
      {...hyperlinkProps}
      className="office-file-xlsx-sheet-grid__floating-chart"
      style={{
        left: XLSX_ROW_HEADER_WIDTH + chart.x,
        top: XLSX_COLUMN_HEADER_HEIGHT + verticalRange.y,
        width: chart.width,
        height: verticalRange.height,
      }}
    >
      <OfficeChartView
        chart={chart.chart}
        width={chart.width}
        height={verticalRange.height}
        zoom={zoom}
      />
    </div>
  );
}

/** 渲染当前范围内一个普通或合并单元格槽位。 */
function VirtualSpreadsheetCell({
  sheetId,
  cell,
  merge,
  left,
  top,
  width,
  height,
  contentBounds,
  viewMode,
  resolveFontFamily,
  frozen = false,
}: {
  sheetId: string;
  cell: SpreadsheetCell;
  merge?: SpreadsheetMerge;
  left: number;
  top: number;
  width: number;
  height: number;
  contentBounds?: SpreadsheetCellContentBounds;
  viewMode: SpreadsheetViewMode;
  /** 当前文档会话统一的字体链解析函数。 */
  resolveFontFamily: OfficeFontFamilyResolver;
  /** 是否属于冻结窗格的重复投影。 */
  frozen?: boolean;
}) {
  const style = cell.style ?? {};
  const shrinkToFit = isSpreadsheetShrinkToFitCell(cell);
  const fallbackBorder = style.border
    ? `${style.borderWidth ?? 1}px solid ${style.borderColor ?? '#b9c2d0'}`
    : '1px solid #d9e0ea';
  const cellStyle: CSSProperties = {
    ...buildXlsxCellStyle(cell, viewMode, resolveFontFamily),
    position: 'absolute',
    left,
    top,
    width,
    height,
    minHeight: height,
    fontSize: style.fontSize ?? (isHighlightedXlsxCell(style) ? 14 : 13),
    borderTop: style.borderTop ?? fallbackBorder,
    borderRight: style.borderRight ?? fallbackBorder,
    borderBottom: style.borderBottom ?? fallbackBorder,
    borderLeft: style.borderLeft ?? fallbackBorder,
    // 有内容的单元格置于空单元格之上，文本才能按 Excel 规则穿过连续空白格。
    zIndex: frozen ? 10 : cell.value ? 2 : 1,
    // 非冻结单元格也必须保留源填充；未填充时继续透明，才能维持 Excel 文本溢出规则。
    background: frozen
      ? style.backgroundColor ?? '#fff'
      : style.backgroundColor,
    boxShadow: frozen ? '1px 1px 0 #91a0b3' : undefined,
  };
  return (
    <div
      className={
        frozen
          ? 'office-file-xlsx-virtual-grid__cell office-file-xlsx-virtual-grid__cell--frozen'
          : 'office-file-xlsx-virtual-grid__cell'
      }
      data-office-spreadsheet-cell={cell.ref}
      tabIndex={-1}
      title={cell.value}
      style={cellStyle}
    >
      <SpreadsheetCellRenderer
        cell={cell}
        sourceId={`${sheetId}:${cell.ref}`}
        sheetId={sheetId}
        contentWidth={width}
        contentHeight={
          viewMode === 'source' || shrinkToFit ? height : undefined
        }
        clipped={Boolean(
          viewMode === 'reading'
            ? shrinkToFit
            : style.wrapText || style.shrinkToFit || merge,
        )}
        contentBounds={viewMode === 'source' ? contentBounds : undefined}
        viewMode={viewMode}
      />
    </div>
  );
}

/** 渲染大型 Sheet 的行列窗口。 */
function VirtualSpreadsheetGridComponent({
  source,
  sheetId,
  layout,
  gridMode,
  zoom,
  viewMode,
  readingRowHeights,
  onReadingRowHeightsChange,
  navigationControllerRef,
}: VirtualSpreadsheetGridProps) {
  const resolveFontFamily = useOfficeFontResolver();
  const {
    viewportRef,
    viewport,
    range,
    sourceRowAxis,
    rowAxis,
    columnAxis,
    scale,
    canvasWidth,
    canvasHeight,
  } = useSpreadsheetGridWindow({
    sheetId,
    layout,
    gridMode,
    zoom,
    viewMode,
    readingRowHeights,
  });
  const generationRef = useRef(0);
  const [retryRevision, setRetryRevision] = useState(0);
  const [loadedRange, setLoadedRange] = useState<LoadedSpreadsheetRange>();
  const [error, setError] = useState<Error>();
  const data =
    loadedRange?.source === source && loadedRange.sheetId === sheetId
      ? loadedRange.data
      : undefined;

  useEffect(() => {
    const controller: SpreadsheetNavigationController = {
      sheetId,
      scrollToCell(rowIndex, columnIndex) {
        const viewportElement = viewportRef.current;
        if (
          !viewportElement ||
          rowIndex < 1 ||
          rowIndex > layout.rowCount ||
          columnIndex < 1 ||
          columnIndex > layout.columnCount
        ) {
          return false;
        }
        const targetLeft =
          (XLSX_ROW_HEADER_WIDTH +
            columnAxis.offsetAt(columnIndex) +
            columnAxis.sizeAt(columnIndex) / 2) *
            scale -
          viewportElement.clientWidth / 2;
        const targetTop =
          (XLSX_COLUMN_HEADER_HEIGHT +
            rowAxis.offsetAt(rowIndex) +
            rowAxis.sizeAt(rowIndex) / 2) *
            scale -
          viewportElement.clientHeight / 2;
        viewportElement.scrollTo({
          left: Math.min(
            Math.max(0, targetLeft),
            Math.max(
              0,
              viewportElement.scrollWidth - viewportElement.clientWidth,
            ),
          ),
          top: Math.min(
            Math.max(0, targetTop),
            Math.max(
              0,
              viewportElement.scrollHeight - viewportElement.clientHeight,
            ),
          ),
          behavior: window.matchMedia?.('(prefers-reduced-motion: reduce)')
            .matches
            ? 'auto'
            : 'smooth',
        });
        return true;
      },
    };
    navigationControllerRef.current = controller;
    return () => {
      if (navigationControllerRef.current === controller) {
        navigationControllerRef.current = undefined;
      }
    };
  }, [
    columnAxis,
    layout.columnCount,
    layout.rowCount,
    navigationControllerRef,
    rowAxis,
    scale,
    sheetId,
    viewportRef,
  ]);

  useIsomorphicLayoutEffect(() => {
    if (viewMode !== 'reading' || !data) return;
    const updates = buildSpreadsheetReadingRowHeightUpdates(
      data,
      rowAxis,
      columnAxis,
    );
    if (updates.size) onReadingRowHeightsChange(sheetId, updates);
  }, [columnAxis, data, onReadingRowHeightsChange, rowAxis, sheetId, viewMode]);

  useEffect(() => {
    const controller = new AbortController();
    const generation = ++generationRef.current;
    let releaseRange: (() => void) | undefined;
    setError(undefined);
    void source.getRange(sheetId, range, controller.signal).then(
      (nextData) => {
        if (controller.signal.aborted || generation !== generationRef.current) {
          return;
        }
        releaseRange = source.retainRange(sheetId, nextData.range);
        // 同一工作表换窗时保留旧数据，待新范围就绪后一次替换，避免滚动闪白。
        setLoadedRange({ source, sheetId, data: nextData });
      },
      (reason) => {
        if (controller.signal.aborted || generation !== generationRef.current) {
          return;
        }
        setLoadedRange((current) =>
          current?.source === source && current.sheetId === sheetId
            ? undefined
            : current,
        );
        setError(
          reason instanceof Error ? reason : new Error('工作表范围加载失败'),
        );
      },
    );
    return () => {
      controller.abort();
      releaseRange?.();
    };
  }, [
    range.endColumn,
    range.endRow,
    range.startColumn,
    range.startRow,
    retryRevision,
    sheetId,
    source,
  ]);

  const cells = useMemo(
    () =>
      new Map(
        data?.cells.map((cell) => [
          positionKey(cell.rowIndex, cell.columnIndex),
          cell,
        ]) ?? [],
      ),
    [data],
  );
  const visibleRows = useMemo(
    () => data?.rows.filter((row) => !row.hidden) ?? [],
    [data],
  );
  const visibleColumns = useMemo(
    () => data?.columns.filter((column) => !column.hidden) ?? [],
    [data],
  );
  const mergeByPosition = useMemo(() => {
    const result = new Map<string, SpreadsheetMerge>();
    data?.merges.forEach((merge) => {
      visibleRows.forEach((row) => {
        if (row.index < merge.startRow || row.index > merge.endRow) return;
        visibleColumns.forEach((column) => {
          if (
            column.index >= merge.startColumn &&
            column.index <= merge.endColumn
          ) {
            result.set(positionKey(row.index, column.index), merge);
          }
        });
      });
    });
    return result;
  }, [data, visibleColumns, visibleRows]);
  const contentBoundsByPosition = useMemo(() => {
    const result = new Map<string, SpreadsheetCellContentBounds>();
    if (!data || viewMode === 'reading') return result;
    const columnWidths = visibleColumns.map((column) =>
      columnAxis.sizeAt(column.index),
    );
    visibleRows.forEach((row) => {
      const occupiedColumns = visibleColumns.map((column) => {
        const key = positionKey(row.index, column.index);
        return Boolean(
          mergeByPosition.has(key) || isSpreadsheetCellOccupied(cells.get(key)),
        );
      });
      const rowCells = visibleColumns.flatMap((column, columnOffset) => {
        const key = positionKey(row.index, column.index);
        const merge = mergeByPosition.get(key);
        if (
          merge &&
          (merge.startRow !== row.index || merge.startColumn !== column.index)
        ) {
          return [];
        }
        const cell = cells.get(key) ?? createEmptyCell(row.index, column.index);
        const columnSpan = merge
          ? visibleColumns.filter(
              (item) =>
                item.index >= merge.startColumn &&
                item.index <= merge.endColumn,
            ).length
          : 1;
        return [
          {
            key,
            cell,
            columnOffset,
            columnSpan,
            clipped: Boolean(
              cell.style?.wrapText || cell.style?.shrinkToFit || merge,
            ),
          },
        ];
      });
      buildSpreadsheetRowContentBounds(
        rowCells,
        occupiedColumns,
        columnWidths,
      ).forEach((bounds, key) => result.set(key, bounds));
    });
    return result;
  }, [
    cells,
    columnAxis,
    data,
    mergeByPosition,
    viewMode,
    visibleColumns,
    visibleRows,
  ]);
  const logicalScrollLeft = viewport.scrollLeft / scale;
  const logicalScrollTop = viewport.scrollTop / scale;
  const headerScrollLeft = Math.max(
    0,
    logicalScrollLeft - XLSX_GRID_VIEWPORT_INSET / scale,
  );
  const headerScrollTop = Math.max(
    0,
    logicalScrollTop - XLSX_GRID_VIEWPORT_INSET / scale,
  );
  const renderFrozenCell = (args: SpreadsheetFrozenCellRenderArgs) => (
    <VirtualSpreadsheetCell
      key={args.key}
      sheetId={sheetId}
      cell={args.cell}
      merge={args.merge}
      left={args.left}
      top={args.top}
      width={args.width}
      height={args.height}
      viewMode={viewMode}
      resolveFontFamily={resolveFontFamily}
      frozen
    />
  );

  return (
    <div
      ref={viewportRef}
      className="office-file-xlsx-sheet-grid office-file-xlsx-virtual-grid"
      data-office-spreadsheet-sheet-id={sheetId}
      data-office-fit-viewport="true"
    >
      <div
        className="office-file-xlsx-sheet-grid__canvas office-file-xlsx-virtual-grid__canvas"
        data-office-fit-target="true"
        data-office-fit-base-width={canvasWidth}
        data-office-fit-base-height={canvasHeight}
        style={{
          width: canvasWidth,
          minWidth: canvasWidth,
          height: canvasHeight,
          minHeight: canvasHeight,
          zoom: scale,
        }}
      >
        {data
          ? visibleRows.flatMap((row) =>
              visibleColumns.flatMap((column) => {
                const key = positionKey(row.index, column.index);
                const merge = mergeByPosition.get(key);
                if (
                  merge &&
                  (merge.startRow !== row.index ||
                    merge.startColumn !== column.index)
                ) {
                  return [];
                }
                const cell =
                  cells.get(key) ?? createEmptyCell(row.index, column.index);
                const endRow = merge?.endRow ?? row.index;
                const endColumn = merge?.endColumn ?? column.index;
                return [
                  <VirtualSpreadsheetCell
                    key={cell.ref}
                    sheetId={sheetId}
                    cell={cell}
                    merge={merge}
                    left={
                      XLSX_ROW_HEADER_WIDTH + columnAxis.offsetAt(column.index)
                    }
                    top={
                      XLSX_COLUMN_HEADER_HEIGHT + rowAxis.offsetAt(row.index)
                    }
                    width={columnAxis.rangeSize(column.index, endColumn)}
                    height={rowAxis.rangeSize(row.index, endRow)}
                    contentBounds={contentBoundsByPosition.get(key)}
                    viewMode={viewMode}
                    resolveFontFamily={resolveFontFamily}
                  />,
                ];
              }),
            )
          : null}

        {layout.pane && layout.pane.state !== 'split' ? (
          <SpreadsheetFrozenPanes
            source={source}
            sheetId={sheetId}
            bodyRange={range}
            rowCount={layout.rowCount}
            columnCount={layout.columnCount}
            frozenRows={layout.pane.frozenRows}
            frozenColumns={layout.pane.frozenColumns}
            logicalScrollLeft={logicalScrollLeft}
            logicalScrollTop={logicalScrollTop}
            logicalViewportWidth={viewport.width / scale}
            logicalViewportHeight={viewport.height / scale}
            rowAxis={rowAxis}
            columnAxis={columnAxis}
            renderCell={renderFrozenCell}
          />
        ) : null}

        {data?.images.map((image) => (
          <VirtualSpreadsheetImage
            key={image.id}
            image={image}
            sourceRowAxis={sourceRowAxis}
            rowAxis={rowAxis}
          />
        ))}
        {data?.charts.map((chart) => (
          <VirtualSpreadsheetChart
            key={chart.id}
            chart={chart}
            zoom={zoom}
            sourceRowAxis={sourceRowAxis}
            rowAxis={rowAxis}
          />
        ))}

        {visibleColumns.map((column) => (
          <div
            key={column.index}
            className="office-file-xlsx-virtual-grid__column-header"
            style={{
              left: XLSX_ROW_HEADER_WIDTH + columnAxis.offsetAt(column.index),
              top: headerScrollTop,
              width: columnAxis.sizeAt(column.index),
              height: XLSX_COLUMN_HEADER_HEIGHT,
            }}
          >
            {getSpreadsheetColumnLabel(column.index)}
          </div>
        ))}
        {visibleRows.map((row) => (
          <div
            key={row.index}
            className="office-file-xlsx-virtual-grid__row-header"
            style={{
              left: headerScrollLeft,
              top: XLSX_COLUMN_HEADER_HEIGHT + rowAxis.offsetAt(row.index),
              width: XLSX_ROW_HEADER_WIDTH,
              height: rowAxis.sizeAt(row.index),
            }}
          >
            {row.index}
          </div>
        ))}
        <div
          className="office-file-xlsx-virtual-grid__corner"
          style={{
            left: headerScrollLeft,
            top: headerScrollTop,
            width: XLSX_ROW_HEADER_WIDTH,
            height: XLSX_COLUMN_HEADER_HEIGHT,
          }}
        />
      </div>
      {!data ? (
        <SpreadsheetGridPlaceholder
          loading={!error}
          error={error}
          onRetry={() => setRetryRevision((value) => value + 1)}
        />
      ) : null}
    </div>
  );
}

function createEmptyCell(
  rowIndex: number,
  columnIndex: number,
): SpreadsheetCell {
  return {
    ref: `${getSpreadsheetColumnLabel(columnIndex)}${rowIndex}`,
    rowIndex,
    columnIndex,
    value: '',
  };
}

export const VirtualSpreadsheetGrid = memo(VirtualSpreadsheetGridComponent);
