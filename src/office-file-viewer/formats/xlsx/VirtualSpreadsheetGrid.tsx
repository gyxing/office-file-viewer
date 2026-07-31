// VirtualSpreadsheetGrid 按完整工作表坐标渲染当前二维窗口。
import type { CSSProperties } from 'react';
import React, { memo, useEffect, useMemo, useRef, useState } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import {
  useOfficeResourceUrl,
  type OfficeResourceSource,
} from '../../services/resource-store';
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
import { OfficeChartView } from '../../shared/chart/OfficeChartView';
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
import { SpreadsheetGridPlaceholder } from './SpreadsheetGridPlaceholder';
import { useSpreadsheetGridWindow } from './useSpreadsheetGridWindow';

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
function VirtualSpreadsheetImage({ image }: { image: SpreadsheetImage }) {
  const messages = useOfficeFileViewerMessages();
  const source = useMemo<OfficeResourceSource>(
    () =>
      typeof image.src === 'string'
        ? { kind: 'url', url: image.src }
        : image.src,
    [image.src],
  );
  const resource = useOfficeResourceUrl(source);
  return (
    <img
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
        top: XLSX_COLUMN_HEADER_HEIGHT + image.y,
        width: image.width,
        height: image.height,
      }}
    />
  );
}

/** 在虚拟画布中按全局锚点渲染单个图表。 */
function VirtualSpreadsheetChart({
  chart,
  zoom,
}: {
  chart: SpreadsheetChart;
  zoom: number;
}) {
  return (
    <div
      className="office-file-xlsx-sheet-grid__floating-chart"
      style={{
        left: XLSX_ROW_HEADER_WIDTH + chart.x,
        top: XLSX_COLUMN_HEADER_HEIGHT + chart.y,
        width: chart.width,
        height: chart.height,
      }}
    >
      <OfficeChartView
        chart={chart.chart}
        width={chart.width}
        height={chart.height}
        zoom={zoom}
      />
    </div>
  );
}

/** 渲染当前范围内一个普通或合并单元格槽位。 */
function VirtualSpreadsheetCell({
  cell,
  merge,
  left,
  top,
  width,
  height,
  contentBounds,
}: {
  cell: SpreadsheetCell;
  merge?: SpreadsheetMerge;
  left: number;
  top: number;
  width: number;
  height: number;
  contentBounds?: SpreadsheetCellContentBounds;
}) {
  const style = cell.style ?? {};
  const fallbackBorder = style.border
    ? `${style.borderWidth ?? 1}px solid ${style.borderColor ?? '#b9c2d0'}`
    : '1px solid #d9e0ea';
  const cellStyle: CSSProperties = {
    ...buildXlsxCellStyle(cell),
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
    zIndex: cell.value ? 2 : 1,
  };
  return (
    <div
      className="office-file-xlsx-virtual-grid__cell"
      title={cell.value}
      style={cellStyle}
    >
      <SpreadsheetCellRenderer
        cell={cell}
        contentWidth={width}
        contentHeight={height}
        clipped={Boolean(style.wrapText || style.shrinkToFit || merge)}
        contentBounds={contentBounds}
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
}: VirtualSpreadsheetGridProps) {
  const {
    viewportRef,
    viewport,
    range,
    rowAxis,
    columnAxis,
    scale,
    canvasWidth,
    canvasHeight,
  } = useSpreadsheetGridWindow(layout, gridMode, zoom);
  const generationRef = useRef(0);
  const [retryRevision, setRetryRevision] = useState(0);
  const [loadedRange, setLoadedRange] = useState<LoadedSpreadsheetRange>();
  const [error, setError] = useState<Error>();
  const data =
    loadedRange?.source === source && loadedRange.sheetId === sheetId
      ? loadedRange.data
      : undefined;

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
    if (!data) return result;
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
  }, [cells, columnAxis, data, mergeByPosition, visibleColumns, visibleRows]);
  const logicalScrollLeft = viewport.scrollLeft / scale;
  const logicalScrollTop = viewport.scrollTop / scale;

  return (
    <div
      ref={viewportRef}
      className="office-file-xlsx-sheet-grid office-file-xlsx-virtual-grid"
    >
      <div
        className="office-file-xlsx-sheet-grid__canvas office-file-xlsx-virtual-grid__canvas"
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
                  />,
                ];
              }),
            )
          : null}

        {data?.images.map((image) => (
          <VirtualSpreadsheetImage key={image.id} image={image} />
        ))}
        {data?.charts.map((chart) => (
          <VirtualSpreadsheetChart key={chart.id} chart={chart} zoom={zoom} />
        ))}

        {visibleColumns.map((column) => (
          <div
            key={column.index}
            className="office-file-xlsx-virtual-grid__column-header"
            style={{
              left: XLSX_ROW_HEADER_WIDTH + columnAxis.offsetAt(column.index),
              top: logicalScrollTop,
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
              left: logicalScrollLeft,
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
            left: logicalScrollLeft,
            top: logicalScrollTop,
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
