import type { ReactNode } from 'react';
import React, { memo, useEffect, useMemo, useRef, useState } from 'react';
import type { SpreadsheetAxisIndex } from '../../services/spreadsheet/SpreadsheetAxisIndex';
import type { SpreadsheetSource } from '../../services/spreadsheet/SpreadsheetSource';
import type {
  SpreadsheetCell,
  SpreadsheetMerge,
  SpreadsheetRange,
  SpreadsheetRangeData,
} from '../../services/spreadsheet/types';
import {
  getSpreadsheetColumnLabel,
  XLSX_COLUMN_HEADER_HEIGHT,
  XLSX_ROW_HEADER_WIDTH,
} from './sheetRenderUtils';

/** 冻结窗格重复投影单元格时交给主虚拟网格的几何参数。 */
export type SpreadsheetFrozenCellRenderArgs = {
  /** 当前槽位对应的单元格。 */
  cell: SpreadsheetCell;
  /** 当前槽位对应的合并区域。 */
  merge?: SpreadsheetMerge;
  /** 画布内左坐标。 */
  left: number;
  /** 画布内顶坐标。 */
  top: number;
  /** 单元格或合并区域宽度。 */
  width: number;
  /** 单元格或合并区域高度。 */
  height: number;
  /** 冻结投影使用的稳定 React 键。 */
  key: string;
};

/** 当前冻结投影加载的三块独立稀疏范围。 */
type FrozenPaneData = {
  /** 冻结行与当前可见列相交的数据。 */
  rows?: SpreadsheetRangeData;
  /** 冻结列与当前可见行相交的数据。 */
  columns?: SpreadsheetRangeData;
  /** 冻结行列交叉角的数据。 */
  corner?: SpreadsheetRangeData;
};

/** 把范围限制在工作表有效边界内。 */
function normalizeRange(
  range: SpreadsheetRange,
  rowCount: number,
  columnCount: number,
) {
  return {
    startRow: Math.max(1, Math.min(rowCount, range.startRow)),
    endRow: Math.max(1, Math.min(rowCount, range.endRow)),
    startColumn: Math.max(1, Math.min(columnCount, range.startColumn)),
    endColumn: Math.max(1, Math.min(columnCount, range.endColumn)),
  };
}

/** 返回范围是否包含至少一个行列位置。 */
function hasRange(range: SpreadsheetRange) {
  return range.endRow >= range.startRow && range.endColumn >= range.startColumn;
}

/** 为范围数据建立合并位置索引。 */
function createMergeIndex(data: SpreadsheetRangeData) {
  const result = new Map<string, SpreadsheetMerge>();
  data.merges.forEach((merge) => {
    for (let row = merge.startRow; row <= merge.endRow; row += 1) {
      for (
        let column = merge.startColumn;
        column <= merge.endColumn;
        column += 1
      ) {
        result.set(`${row}:${column}`, merge);
      }
    }
  });
  return result;
}

/** 渲染大型工作表的冻结角、冻结行和冻结列投影视图。 */
function SpreadsheetFrozenPanesComponent({
  source,
  sheetId,
  bodyRange,
  rowCount,
  columnCount,
  frozenRows,
  frozenColumns,
  logicalScrollLeft,
  logicalScrollTop,
  logicalViewportWidth,
  logicalViewportHeight,
  rowAxis,
  columnAxis,
  renderCell,
}: {
  /** 当前工作簿按需数据源。 */
  source: SpreadsheetSource;
  /** 当前工作表稳定标识。 */
  sheetId: string;
  /** 主体窗口已经请求的行列范围。 */
  bodyRange: SpreadsheetRange;
  /** 工作表总行数。 */
  rowCount: number;
  /** 工作表总列数。 */
  columnCount: number;
  /** 顶部冻结行数量。 */
  frozenRows: number;
  /** 左侧冻结列数量。 */
  frozenColumns: number;
  /** 未缩放的水平滚动位置。 */
  logicalScrollLeft: number;
  /** 未缩放的垂直滚动位置。 */
  logicalScrollTop: number;
  /** 未缩放的视口宽度。 */
  logicalViewportWidth: number;
  /** 未缩放的视口高度。 */
  logicalViewportHeight: number;
  /** 当前显示模式使用的行坐标索引。 */
  rowAxis: SpreadsheetAxisIndex;
  /** 当前显示模式使用的列坐标索引。 */
  columnAxis: SpreadsheetAxisIndex;
  /** 使用主虚拟网格单元格渲染器绘制重复投影。 */
  renderCell(args: SpreadsheetFrozenCellRenderArgs): ReactNode;
}) {
  const generationRef = useRef(0);
  const [data, setData] = useState<FrozenPaneData>({});
  const visibleFrozenRows = Math.min(
    frozenRows,
    rowAxis.findIndexAtOffset(Math.max(0, logicalViewportHeight)),
  );
  const visibleFrozenColumns = Math.min(
    frozenColumns,
    columnAxis.findIndexAtOffset(Math.max(0, logicalViewportWidth)),
  );
  const ranges = useMemo(() => {
    const rows = normalizeRange(
      {
        startRow: 1,
        endRow: visibleFrozenRows,
        startColumn: Math.max(1, bodyRange.startColumn),
        endColumn: bodyRange.endColumn,
      },
      rowCount,
      columnCount,
    );
    const columns = normalizeRange(
      {
        startRow: Math.max(1, bodyRange.startRow),
        endRow: bodyRange.endRow,
        startColumn: 1,
        endColumn: visibleFrozenColumns,
      },
      rowCount,
      columnCount,
    );
    const corner = normalizeRange(
      {
        startRow: 1,
        endRow: visibleFrozenRows,
        startColumn: 1,
        endColumn: visibleFrozenColumns,
      },
      rowCount,
      columnCount,
    );
    return { rows, columns, corner };
  }, [
    bodyRange,
    columnCount,
    rowCount,
    visibleFrozenColumns,
    visibleFrozenRows,
  ]);

  useEffect(() => {
    const controller = new AbortController();
    const generation = ++generationRef.current;
    const entries = (
      [
        ['rows', ranges.rows, visibleFrozenRows > 0],
        ['columns', ranges.columns, visibleFrozenColumns > 0],
        [
          'corner',
          ranges.corner,
          visibleFrozenRows > 0 && visibleFrozenColumns > 0,
        ],
      ] as const
    ).filter(([, range, enabled]) => enabled && hasRange(range));
    const releases: Array<() => void> = [];
    void Promise.all(
      entries.map(async ([kind, range]) => {
        const next = await source.getRange(sheetId, range, controller.signal);
        releases.push(source.retainRange(sheetId, next.range));
        return [kind, next] as const;
      }),
    ).then(
      (results) => {
        if (controller.signal.aborted || generation !== generationRef.current)
          return;
        setData(Object.fromEntries(results));
      },
      () => undefined,
    );
    return () => {
      controller.abort();
      releases.forEach((release) => release());
    };
  }, [ranges, sheetId, source, visibleFrozenColumns, visibleFrozenRows]);

  const renderRange = (
    rangeData: SpreadsheetRangeData | undefined,
    kind: 'rows' | 'columns' | 'corner',
  ) => {
    if (!rangeData) return null;
    const cellByPosition = new Map(
      rangeData.cells.map((cell) => [
        `${cell.rowIndex}:${cell.columnIndex}`,
        cell,
      ]),
    );
    const merges = createMergeIndex(rangeData);
    const rows = rangeData.rows.filter((row) => !row.hidden);
    const columns = rangeData.columns.filter((column) => !column.hidden);
    return rows.flatMap((row) =>
      columns.flatMap((column) => {
        const key = `${row.index}:${column.index}`;
        const merge = merges.get(key);
        if (
          merge &&
          (merge.startRow !== row.index || merge.startColumn !== column.index)
        ) {
          return [];
        }
        const cell = cellByPosition.get(key) ?? {
          ref: `${getSpreadsheetColumnLabel(column.index)}${row.index}`,
          rowIndex: row.index,
          columnIndex: column.index,
          value: '',
        };
        const left =
          XLSX_ROW_HEADER_WIDTH +
          columnAxis.offsetAt(column.index) +
          (kind === 'columns' || kind === 'corner' ? logicalScrollLeft : 0);
        const top =
          XLSX_COLUMN_HEADER_HEIGHT +
          rowAxis.offsetAt(row.index) +
          (kind === 'rows' || kind === 'corner' ? logicalScrollTop : 0);
        return [
          renderCell({
            key: `frozen-${kind}-${cell.ref}`,
            cell,
            merge,
            left,
            top,
            width: columnAxis.rangeSize(
              column.index,
              merge?.endColumn ?? column.index,
            ),
            height: rowAxis.rangeSize(row.index, merge?.endRow ?? row.index),
          }),
        ];
      }),
    );
  };

  if (!visibleFrozenRows && !visibleFrozenColumns) return null;
  return (
    <div className="office-file-xlsx-frozen-panes">
      {renderRange(data.rows, 'rows')}
      {renderRange(data.columns, 'columns')}
      {renderRange(data.corner, 'corner')}
      {visibleFrozenColumns
        ? data.corner?.columns
            .filter((column) => !column.hidden)
            .map((column) => (
              <div
                key={`frozen-column-header-${column.index}`}
                className="office-file-xlsx-virtual-grid__column-header office-file-xlsx-frozen-panes__header"
                style={{
                  left:
                    logicalScrollLeft +
                    XLSX_ROW_HEADER_WIDTH +
                    columnAxis.offsetAt(column.index),
                  top: logicalScrollTop,
                  width: columnAxis.sizeAt(column.index),
                  height: XLSX_COLUMN_HEADER_HEIGHT,
                }}
              >
                {getSpreadsheetColumnLabel(column.index)}
              </div>
            ))
        : null}
      {visibleFrozenRows
        ? data.corner?.rows
            .filter((row) => !row.hidden)
            .map((row) => (
              <div
                key={`frozen-row-header-${row.index}`}
                className="office-file-xlsx-virtual-grid__row-header office-file-xlsx-frozen-panes__header"
                style={{
                  left: logicalScrollLeft,
                  top:
                    logicalScrollTop +
                    XLSX_COLUMN_HEADER_HEIGHT +
                    rowAxis.offsetAt(row.index),
                  width: XLSX_ROW_HEADER_WIDTH,
                  height: rowAxis.sizeAt(row.index),
                }}
              >
                {row.index}
              </div>
            ))
        : null}
      {visibleFrozenRows && visibleFrozenColumns ? (
        <div
          className="office-file-xlsx-virtual-grid__corner office-file-xlsx-frozen-panes__header"
          style={{
            left: logicalScrollLeft,
            top: logicalScrollTop,
            width: XLSX_ROW_HEADER_WIDTH,
            height: XLSX_COLUMN_HEADER_HEIGHT,
          }}
        />
      ) : null}
    </div>
  );
}

export const SpreadsheetFrozenPanes = memo(SpreadsheetFrozenPanesComponent);
