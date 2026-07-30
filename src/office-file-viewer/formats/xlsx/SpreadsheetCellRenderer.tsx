// SpreadsheetCellRenderer 统一渲染普通表格与窗口表格的单元格内容和对角边框。
import React, { memo } from 'react';
import type {
  SpreadsheetCell,
  SpreadsheetDiagonalBorder,
} from '../../services/spreadsheet/types';

/** 定义 SpreadsheetCellRenderer 组件可接收的属性。 */
type SpreadsheetCellRendererProps = {
  /** 当前单元格。 */
  cell: SpreadsheetCell;
  /** 单元格内容允许使用的最大高度。 */
  contentHeight?: number;
  /** 是否按换行或合并单元格规则裁切内容。 */
  clipped?: boolean;
};

/** 根据源文件线型生成浏览器 SVG 描边间隔。 */
function getStrokeDashArray(lineStyle: SpreadsheetDiagonalBorder['lineStyle']) {
  if (lineStyle === 'dotted' || lineStyle === 'hair') return '1 2';
  if (lineStyle === 'dashed') return '6 3';
  if (lineStyle === 'dashDot') return '7 3 1 3';
  if (lineStyle === 'dashDotDot' || lineStyle === 'slantDashDot') {
    return '7 3 1 3 1 3';
  }
  return undefined;
}

/** 渲染一个严格贴合最终单元格边框盒的对角边框层。 */
function SpreadsheetDiagonalBorderLayer({
  border,
}: {
  border: SpreadsheetDiagonalBorder;
}) {
  const dashArray = getStrokeDashArray(border.lineStyle);
  const common = {
    stroke: border.color,
    strokeWidth: border.width,
    strokeDasharray: dashArray,
    vectorEffect: 'non-scaling-stroke' as const,
  };
  return (
    <svg
      aria-hidden="true"
      className="office-file-xlsx-cell-diagonal"
      preserveAspectRatio="none"
      viewBox="0 0 100 100"
    >
      {border.down ? (
        <line x1="0" y1="0" x2="100" y2="100" {...common} />
      ) : null}
      {border.up ? <line x1="0" y1="100" x2="100" y2="0" {...common} /> : null}
      {border.lineStyle === 'double' ? (
        <>
          {border.down ? (
            <line x1="0" y1="2" x2="98" y2="100" {...common} strokeWidth={1} />
          ) : null}
          {border.up ? (
            <line x1="0" y1="98" x2="98" y2="0" {...common} strokeWidth={1} />
          ) : null}
        </>
      ) : null}
    </svg>
  );
}

/** 渲染电子表格单元格的共享内容层。 */
function SpreadsheetCellRendererComponent({
  cell,
  contentHeight,
  clipped,
}: SpreadsheetCellRendererProps) {
  const clippedPosition = clipped
    ? cell.style?.verticalAlign === 'bottom'
      ? { bottom: 0 }
      : cell.style?.verticalAlign === 'middle'
      ? { top: '50%', transform: 'translateY(-50%)' }
      : { top: 0 }
    : undefined;
  return (
    <>
      <div
        className={`office-file-xlsx-sheet-table__cell-content${
          clipped ? ' office-file-xlsx-sheet-table__cell-content--clipped' : ''
        }`}
        style={{ maxHeight: contentHeight, ...clippedPosition }}
      >
        {cell.value}
      </div>
      {cell.style?.diagonalBorder ? (
        <SpreadsheetDiagonalBorderLayer border={cell.style.diagonalBorder} />
      ) : null}
    </>
  );
}

export const SpreadsheetCellRenderer = memo(SpreadsheetCellRendererComponent);
