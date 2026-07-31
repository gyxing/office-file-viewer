// XlsxSheetTable 将工作表行列和单元格模型渲染为带表头的 HTML 表格。
import type { CSSProperties, RefObject } from 'react';
import React, { memo, useMemo } from 'react';
import type { XlsxSheet } from '../../services/xlsx/types';
import { buildXlsxCellStyle, isHighlightedXlsxCell } from './sheetRenderUtils';
import { buildXlsxVisibleTableModel } from './sheetTableUtils';
import { buildSpreadsheetRowContentBounds } from './spreadsheetCellOverflow';
import { SpreadsheetCellRenderer } from './SpreadsheetCellRenderer';

/** Excel工作表表格组件属性。 */
type XlsxSheetTableProps = {
  /** 当前处理的工作表。 */
  sheet: XlsxSheet;
  /** 过滤隐藏列后的真实表格宽度，单位为标准化像素。 */
  tableWidth: number;
  /** 过滤隐藏列后的标准化列宽集合。 */
  visibleColumnWidths: number[];
  /** 过滤隐藏行后的标准化行高集合。 */
  visibleRowHeights: number[];
  /** 供布局层读取浏览器真实表格尺寸的节点引用。 */
  tableRef: RefObject<HTMLTableElement>;
};

/** 使用普通表格模式渲染小型工作表。 */
function XlsxSheetTableComponent({
  sheet,
  tableWidth,
  visibleColumnWidths,
  visibleRowHeights,
  tableRef,
}: XlsxSheetTableProps) {
  const tableModel = useMemo(
    () =>
      buildXlsxVisibleTableModel(sheet, visibleColumnWidths, visibleRowHeights),
    [sheet, visibleColumnWidths, visibleRowHeights],
  );
  const cellStyleCache = useMemo(() => {
    const cache = new Map<string, CSSProperties>();
    tableModel.rows.forEach((row) => {
      row.cells.forEach(({ cell }) => {
        const important = isHighlightedXlsxCell(cell.style);
        // 大表格渲染时单元格很多，先按 ref 缓存静态样式，避免每次 JSX 展开都重复计算。
        cache.set(cell.ref, {
          fontSize: important ? 14 : 13,
          ...buildXlsxCellStyle(cell),
        });
      });
    });
    return cache;
  }, [tableModel.rows]);
  const rowContentBounds = useMemo(
    () =>
      tableModel.rows.map((row) =>
        buildSpreadsheetRowContentBounds(
          row.cells.map(({ cell, columnOffset, colSpan, rowSpan }) => ({
            key: cell.ref,
            cell,
            columnOffset,
            columnSpan: colSpan,
            clipped: Boolean(
              cell.style?.wrapText ||
                cell.style?.shrinkToFit ||
                colSpan ||
                rowSpan,
            ),
          })),
          row.occupiedColumns,
          tableModel.columns.map(({ width }) => width),
        ),
      ),
    [tableModel.columns, tableModel.rows],
  );

  return (
    <table
      ref={tableRef}
      className="office-file-xlsx-sheet-table"
      style={{ width: tableWidth }}
    >
      <colgroup>
        <col className="office-file-xlsx-sheet-table__row-header-col" />
        {tableModel.columns.map(({ column, width }) => (
          <col key={column.index} style={{ width }} />
        ))}
      </colgroup>
      <thead>
        <tr>
          <th className="office-file-xlsx-sheet-table__corner" />
          {tableModel.columns.map(({ column }) => (
            <th
              key={column.index}
              className="office-file-xlsx-sheet-table__column-header"
            >
              {column.label}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {tableModel.rows.map(({ row, height, cells }, rowOffset) => (
          <tr key={row.index} style={{ height }}>
            <th className="office-file-xlsx-sheet-table__row-header">
              <span className="office-file-xlsx-sheet-table__row-header-label">
                {row.index}
              </span>
            </th>
            {cells.map(({ cell, columnOffset, colSpan, rowSpan }) => {
              const style = cell.style ?? {};
              const merged = Boolean(colSpan || rowSpan);
              // 自动行高允许换行文字自然撑开；手动行高和合并区域必须保持源文件边界。
              const clipped = Boolean(
                merged ||
                  style.shrinkToFit ||
                  (style.wrapText && row.customHeight),
              );
              const contentWidth = tableModel.columns
                .slice(columnOffset, columnOffset + (colSpan ?? 1))
                .reduce((sum, item) => sum + item.width, 0);
              // 合并单元格按其跨越行的总高度裁切，避免文本反向撑大源文件行高。
              const contentHeight = tableModel.rows
                .slice(rowOffset, rowOffset + (rowSpan ?? 1))
                .reduce((sum, item) => sum + item.height, 0);
              const cellHeight = rowSpan ? contentHeight : height;
              const fallbackBorder = style.border
                ? `${style.borderWidth ?? 1}px solid ${
                    style.borderColor ?? '#b9c2d0'
                  }`
                : '1px solid #d9e0ea';
              return (
                <td
                  key={cell.ref}
                  className="office-file-xlsx-sheet-table__cell"
                  colSpan={colSpan}
                  rowSpan={rowSpan}
                  title={cell.value}
                  style={{
                    ...cellStyleCache.get(cell.ref),
                    height: cellHeight,
                    minHeight: cellHeight,
                    borderTop: style.borderTop ?? fallbackBorder,
                    borderRight: style.borderRight ?? fallbackBorder,
                    borderBottom: style.borderBottom ?? fallbackBorder,
                    borderLeft: style.borderLeft ?? fallbackBorder,
                  }}
                >
                  <SpreadsheetCellRenderer
                    cell={cell}
                    contentWidth={contentWidth}
                    contentHeight={clipped ? contentHeight : undefined}
                    clipped={clipped}
                    contentBounds={rowContentBounds[rowOffset].get(cell.ref)}
                  />
                </td>
              );
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

export const XlsxSheetTable = memo(XlsxSheetTableComponent);
