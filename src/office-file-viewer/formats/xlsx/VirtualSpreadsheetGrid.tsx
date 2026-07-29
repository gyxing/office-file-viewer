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
import { SpreadsheetCellRenderer } from './SpreadsheetCellRenderer';
import { SpreadsheetGridPlaceholder } from './SpreadsheetGridPlaceholder';
import { useSpreadsheetGridWindow } from './useSpreadsheetGridWindow';

type VirtualSpreadsheetGridProps = {
  source: SpreadsheetSource;
  sheetId: string;
  layout: SpreadsheetSheetLayout;
  gridMode: SpreadsheetPerformanceProfile['gridMode'];
  zoom: number;
};

function positionKey(row: number, column: number) {
  return `${row}:${column}`;
}

function findMerge(
  merges: readonly SpreadsheetMerge[],
  row: number,
  column: number,
) {
  return merges.find(
    (merge) =>
      row >= merge.startRow &&
      row <= merge.endRow &&
      column >= merge.startColumn &&
      column <= merge.endColumn,
  );
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
}: {
  cell: SpreadsheetCell;
  merge?: SpreadsheetMerge;
  left: number;
  top: number;
  width: number;
  height: number;
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
    fontSize: isHighlightedXlsxCell(style) ? 14 : 13,
    borderTop: style.borderTop ?? fallbackBorder,
    borderRight: style.borderRight ?? fallbackBorder,
    borderBottom: style.borderBottom ?? fallbackBorder,
    borderLeft: style.borderLeft ?? fallbackBorder,
  };
  return (
    <div
      className="office-file-xlsx-virtual-grid__cell"
      title={cell.value}
      style={cellStyle}
    >
      <SpreadsheetCellRenderer
        cell={cell}
        contentHeight={height}
        clipped={Boolean(style.wrapText || merge)}
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
  const [data, setData] = useState<SpreadsheetRangeData>();
  const [error, setError] = useState<Error>();

  useEffect(() => {
    const controller = new AbortController();
    const generation = ++generationRef.current;
    let releaseRange: (() => void) | undefined;
    setData(undefined);
    setError(undefined);
    void source.getRange(sheetId, range, controller.signal).then(
      (nextData) => {
        if (controller.signal.aborted || generation !== generationRef.current) {
          return;
        }
        releaseRange = source.retainRange(sheetId, nextData.range);
        setData(nextData);
      },
      (reason) => {
        if (controller.signal.aborted || generation !== generationRef.current) {
          return;
        }
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
  const visibleRows = data?.rows.filter((row) => !row.hidden) ?? [];
  const visibleColumns = data?.columns.filter((column) => !column.hidden) ?? [];
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
                const merge = findMerge(data.merges, row.index, column.index);
                if (
                  merge &&
                  (merge.startRow !== row.index ||
                    merge.startColumn !== column.index)
                ) {
                  return [];
                }
                const cell =
                  cells.get(positionKey(row.index, column.index)) ??
                  createEmptyCell(row.index, column.index);
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

function createEmptyCell(rowIndex: number, columnIndex: number) {
  return {
    ref: `${getSpreadsheetColumnLabel(columnIndex)}${rowIndex}`,
    rowIndex,
    columnIndex,
    value: '',
  };
}

export const VirtualSpreadsheetGrid = memo(VirtualSpreadsheetGridComponent);
