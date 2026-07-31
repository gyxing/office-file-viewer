// sheetTableUtils 将含隐藏行列和合并区域的工作表投影为可见表格模型。
import type {
  XlsxCell,
  XlsxColumn,
  XlsxRow,
  XlsxSheet,
} from '../../services/xlsx/types';
import { isSpreadsheetCellOccupied } from './spreadsheetCellOverflow';

/** 描述渲染层实际输出的一列。 */
export type XlsxVisibleTableColumn = {
  /** 对应的源工作表列。 */
  column: XlsxColumn;
  /** 过滤隐藏列后使用的标准化宽度。 */
  width: number;
};

/** 描述渲染层实际输出的一个单元格。 */
export type XlsxVisibleTableCell = {
  /** 单元格内容和样式的来源模型。 */
  cell: XlsxCell;
  /** 当前单元格在可见列集合中的起始位置。 */
  columnOffset: number;
  /** 过滤隐藏列后仍需跨越的可见列数。 */
  colSpan?: number;
  /** 过滤隐藏行后仍需跨越的可见行数。 */
  rowSpan?: number;
};

/** 描述渲染层实际输出的一行。 */
export type XlsxVisibleTableRow = {
  /** 对应的源工作表行。 */
  row: XlsxRow;
  /** 过滤隐藏行后使用的标准化高度。 */
  height: number;
  /** 该行经过隐藏行列和合并区域投影后的单元格。 */
  cells: XlsxVisibleTableCell[];
  /** 标记会截停相邻非换行文本的可见列。 */
  occupiedColumns: boolean[];
};

/** 描述工作表过滤隐藏行列后的完整表格模型。 */
export type XlsxVisibleTableModel = {
  /** 按源顺序保留的可见列。 */
  columns: XlsxVisibleTableColumn[];
  /** 按源顺序保留的可见行。 */
  rows: XlsxVisibleTableRow[];
};

/** 描述一个合并区域投影到可见行列后的渲染信息。 */
type VisibleMergeProjection = {
  /** 合并区域左上角单元格提供的内容和样式。 */
  cell: XlsxCell;
  /** 投影后承载合并单元格的首个可见行索引。 */
  representativeRowIndex: number;
  /** 投影后承载合并单元格的首个可见列索引。 */
  representativeColumnIndex: number;
  /** 投影后横向跨越的可见列数。 */
  colSpan: number;
  /** 投影后纵向跨越的可见行数。 */
  rowSpan: number;
};

/** 生成行列位置使用的稳定键，避免依赖单元格是否实际存在。 */
function positionKey(rowIndex: number, columnIndex: number) {
  return `${rowIndex}:${columnIndex}`;
}

/** 为异常缺失的矩阵位置创建空单元格，保证后续列不会向左错位。 */
function createEmptyCell(rowIndex: number, column: XlsxColumn): XlsxCell {
  return {
    ref: `${column.label}${rowIndex}`,
    rowIndex,
    columnIndex: column.index,
    value: '',
  };
}

/**
 * 将工作表投影为仅包含可见行列的表格模型。
 *
 * 合并区域会重新计算可见跨度；即使原合并锚点落在隐藏行列中，也会把内容迁移到
 * 区域内首个可见位置，避免表格结构和内容同时丢失。
 */
export function buildXlsxVisibleTableModel(
  sheet: XlsxSheet,
  visibleColumnWidths: number[],
  visibleRowHeights: number[],
): XlsxVisibleTableModel {
  const columns = sheet.columns
    .filter((column) => !column.hidden)
    .map((column, index) => ({
      column,
      width: visibleColumnWidths[index],
    }));
  const sourceRows = sheet.rows.filter((row) => !row.hidden);
  const cellByPosition = new Map<string, XlsxCell>();
  sheet.rows.forEach((row) => {
    row.cells.forEach((cell) => {
      cellByPosition.set(positionKey(cell.rowIndex, cell.columnIndex), cell);
    });
  });

  const mergeByPosition = new Map<string, VisibleMergeProjection>();
  sheet.merges.forEach((merge) => {
    const mergeColumns = columns.filter(
      ({ column }) =>
        column.index >= merge.startColumn && column.index <= merge.endColumn,
    );
    const mergeRows = sourceRows.filter(
      (row) => row.index >= merge.startRow && row.index <= merge.endRow,
    );
    if (!mergeColumns.length || !mergeRows.length) return;

    const rootCell =
      cellByPosition.get(positionKey(merge.startRow, merge.startColumn)) ??
      createEmptyCell(merge.startRow, {
        index: merge.startColumn,
        label: merge.ref.split(':')[0].replace(/\d+$/, ''),
        width: mergeColumns[0].width,
      });
    const projection: VisibleMergeProjection = {
      cell: rootCell,
      representativeRowIndex: mergeRows[0].index,
      representativeColumnIndex: mergeColumns[0].column.index,
      colSpan: mergeColumns.length,
      rowSpan: mergeRows.length,
    };
    mergeRows.forEach((row) => {
      mergeColumns.forEach(({ column }) => {
        mergeByPosition.set(positionKey(row.index, column.index), projection);
      });
    });
  });

  const rows = sourceRows.map((row, rowOffset) => {
    const occupiedColumns = columns.map(({ column }) => {
      const key = positionKey(row.index, column.index);
      return Boolean(
        mergeByPosition.has(key) ||
          isSpreadsheetCellOccupied(cellByPosition.get(key)),
      );
    });
    const cells = columns.flatMap<XlsxVisibleTableCell>(
      ({ column }, columnOffset) => {
        const key = positionKey(row.index, column.index);
        const merge = mergeByPosition.get(key);
        if (merge) {
          if (
            row.index !== merge.representativeRowIndex ||
            column.index !== merge.representativeColumnIndex
          ) {
            return [];
          }
          return [
            {
              cell: merge.cell,
              columnOffset,
              colSpan: merge.colSpan > 1 ? merge.colSpan : undefined,
              rowSpan: merge.rowSpan > 1 ? merge.rowSpan : undefined,
            },
          ];
        }

        const cell =
          cellByPosition.get(key) ?? createEmptyCell(row.index, column);
        return [{ cell, columnOffset }];
      },
    );
    return {
      row,
      height: visibleRowHeights[rowOffset],
      cells,
      occupiedColumns,
    };
  });

  return { columns, rows };
}
