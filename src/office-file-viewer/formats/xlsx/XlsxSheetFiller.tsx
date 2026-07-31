// XlsxSheetFiller 使用线性数量的节点补齐工作表可视空白区，避免生成二维单元格矩阵。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { XlsxSheet } from '../../services/xlsx/types';
import type { XlsxSheetMetrics } from './sheetRenderUtils';
import {
  getSpreadsheetColumnLabel,
  XLSX_ROW_HEADER_WIDTH,
} from './sheetRenderUtils';

/** Excel工作表填充层组件属性。 */
type XlsxSheetFillerProps = {
  /** 当前需要补齐可视空白区的工作表。 */
  sheet: XlsxSheet;
  /** 当前视口、缩放比例和工作表共同计算出的渲染尺寸。 */
  metrics: XlsxSheetMetrics;
};

/** 将一组尺寸转换为每个区域末端相对起点的累计偏移。 */
function buildBoundaryOffsets(sizes: number[]) {
  let offset = 0;
  return sizes.map((size) => {
    offset += size;
    return offset;
  });
}

/** 渲染不承载数据的工作表空白行列。 */
function XlsxSheetFillerComponent({ sheet, metrics }: XlsxSheetFillerProps) {
  const fillerWidth = Math.max(
    0,
    metrics.renderedTableWidth - metrics.tableWidth,
  );
  const fillerHeight = Math.max(
    0,
    metrics.renderedTableHeight - metrics.tableHeight,
  );
  const sourceBodyWidth = Math.max(
    0,
    metrics.tableWidth - XLSX_ROW_HEADER_WIDTH,
  );
  const sourceBodyHeight = Math.max(
    0,
    metrics.tableHeight - metrics.columnHeaderHeight,
  );
  const fillerColumns = useMemo(
    () =>
      Array.from({ length: metrics.fillerColumnCount }, (_, offset) => {
        const index = sheet.columnCount + offset + 1;
        return { index, label: getSpreadsheetColumnLabel(index) };
      }),
    [metrics.fillerColumnCount, sheet.columnCount],
  );
  const fillerRows = useMemo(
    () =>
      Array.from(
        { length: metrics.fillerRowCount },
        (_, offset) => sheet.rowCount + offset + 1,
      ),
    [metrics.fillerRowCount, sheet.rowCount],
  );
  const rowBoundaries = useMemo(
    () => buildBoundaryOffsets(metrics.visibleRowHeights),
    [metrics.visibleRowHeights],
  );
  const columnBoundaries = useMemo(
    () => buildBoundaryOffsets(metrics.visibleColumnWidths),
    [metrics.visibleColumnWidths],
  );
  const verticalGridStyle: CSSProperties = {
    backgroundSize: `${metrics.defaultColumnWidth}px 100%`,
  };
  const horizontalGridStyle: CSSProperties = {
    backgroundSize: `100% ${metrics.defaultRowHeight}px`,
  };
  const fullGridStyle: CSSProperties = {
    backgroundSize: `${metrics.defaultColumnWidth}px ${metrics.defaultRowHeight}px`,
  };

  return (
    <div
      className="office-file-xlsx-sheet-filler"
      style={{
        width: metrics.renderedTableWidth,
        height: metrics.renderedTableHeight,
      }}
      aria-hidden="true"
    >
      {fillerColumns.length ? (
        <div
          className="office-file-xlsx-sheet-filler__right"
          style={{
            left: metrics.tableWidth,
            width: fillerWidth,
            height: metrics.renderedTableHeight,
          }}
        >
          <div
            className="office-file-xlsx-sheet-filler__column-headers"
            style={{ height: metrics.columnHeaderHeight }}
          >
            <div className="office-file-xlsx-sheet-filler__column-header-clip">
              {fillerColumns.map((column) => (
                <div
                  key={column.index}
                  className="office-file-xlsx-sheet-filler__column-header"
                  style={{
                    width: metrics.defaultColumnWidth,
                    height: metrics.columnHeaderHeight,
                  }}
                >
                  {column.label}
                </div>
              ))}
            </div>
          </div>
          <div
            className="office-file-xlsx-sheet-filler__source-rows"
            style={{
              ...verticalGridStyle,
              top: metrics.columnHeaderHeight,
              width: fillerWidth,
              height: sourceBodyHeight,
            }}
          >
            {rowBoundaries.map((offset, index) => (
              <i
                key={`${index}:${offset}`}
                className="office-file-xlsx-sheet-filler__row-boundary"
                style={{ top: offset }}
              />
            ))}
          </div>
          {fillerRows.length ? (
            <div
              className="office-file-xlsx-sheet-filler__full-grid"
              style={{
                ...fullGridStyle,
                top: metrics.tableHeight,
                width: fillerWidth,
                height: fillerHeight,
              }}
            />
          ) : null}
        </div>
      ) : null}
      {fillerRows.length ? (
        <div
          className="office-file-xlsx-sheet-filler__bottom"
          style={{
            top: metrics.tableHeight,
            width: metrics.renderedTableWidth,
            height: fillerHeight,
          }}
        >
          <div className="office-file-xlsx-sheet-filler__row-headers">
            <div className="office-file-xlsx-sheet-filler__row-header-clip">
              {fillerRows.map((rowIndex) => (
                <div
                  key={rowIndex}
                  className="office-file-xlsx-sheet-filler__row-header"
                  style={{ height: metrics.defaultRowHeight }}
                >
                  {rowIndex}
                </div>
              ))}
            </div>
          </div>
          <div
            className="office-file-xlsx-sheet-filler__source-columns"
            style={{
              ...horizontalGridStyle,
              left: XLSX_ROW_HEADER_WIDTH,
              width: sourceBodyWidth,
              height: fillerHeight,
            }}
          >
            {columnBoundaries.map((offset, index) => (
              <i
                key={`${index}:${offset}`}
                className="office-file-xlsx-sheet-filler__column-boundary"
                style={{ left: offset }}
              />
            ))}
          </div>
        </div>
      ) : null}
    </div>
  );
}

export const XlsxSheetFiller = memo(XlsxSheetFillerComponent);
