// SpreadsheetCellRenderer 统一渲染普通表格与窗口表格的单元格内容和对角边框。
import React, { memo } from 'react';
import type {
  SpreadsheetCell,
  SpreadsheetDiagonalBorder,
} from '../../services/spreadsheet/types';
import type { SpreadsheetViewMode } from '../../services/spreadsheet/viewMode';
import { useOfficeHyperlink } from '../../shared/hyperlink';
import { OfficeSearchHighlightedText } from '../search/OfficeSearchContext';
import { SpreadsheetAnnotationMarker } from './SpreadsheetAnnotationMarker';
import type { SpreadsheetCellContentBounds } from './spreadsheetCellOverflow';
import {
  estimateSpreadsheetLineWidth,
  estimateSpreadsheetWrappedLineCount,
  isSpreadsheetShrinkToFitCell,
  resolveSpreadsheetCellFontSize,
  SPREADSHEET_CONTENT_HORIZONTAL_INSET,
  SPREADSHEET_CONTENT_LINE_HEIGHT,
  SPREADSHEET_CONTENT_VERTICAL_INSET,
  SPREADSHEET_CONTENT_WIDTH_SAFETY_RATIO,
} from './spreadsheetReadingLayout';

/** 电子表格单元格渲染器组件属性。 */
type SpreadsheetCellRendererProps = {
  /** 当前单元格。 */
  cell: SpreadsheetCell;
  /** 单元格内容区宽度，用于还原缩小字体以适应。 */
  contentWidth?: number;
  /** 单元格内容允许使用的最大高度。 */
  contentHeight?: number;
  /** 是否按换行或合并单元格规则裁切内容。 */
  clipped?: boolean;
  /** 非换行文本遇到相邻阻挡单元格时使用的可见走廊。 */
  contentBounds?: SpreadsheetCellContentBounds;
  /** 当前电子表格采用的显示模式。 */
  viewMode?: SpreadsheetViewMode;
  /** 当前单元格在工作簿中的稳定来源标识。 */
  sourceId: string;
  /** 当前单元格所属工作表的稳定标识。 */
  sheetId: string;
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

/** 按单元格宽度计算 shrink-to-fit 字号，不触发布局读取。 */
function resolveShrinkToFitFontSize(
  cell: SpreadsheetCell,
  contentWidth: number | undefined,
) {
  const style = cell.style;
  if (!style?.shrinkToFit || style.wrapText || !contentWidth || !cell.value) {
    return undefined;
  }
  const fontSize = resolveSpreadsheetCellFontSize(cell);
  const textWidth = Math.max(
    ...cell.value
      .split(/\r?\n/)
      .map((line) => estimateSpreadsheetLineWidth(line, fontSize)),
  );
  const availableWidth = Math.max(1, contentWidth - 4);
  if (textWidth <= availableWidth) return undefined;
  return Math.max(
    1,
    Math.floor((fontSize * availableWidth * 100) / textWidth) / 100,
  );
}

/** 为固定边界内的换行内容预先计算最大可容纳字号，避免渲染后裁切或撑破单元格。 */
function resolveBoundedWrapFontSize(
  cell: SpreadsheetCell,
  contentWidth: number | undefined,
  contentHeight: number | undefined,
  clipped: boolean | undefined,
) {
  const style = cell.style;
  if (
    !clipped ||
    !style?.wrapText ||
    !contentWidth ||
    !contentHeight ||
    !cell.value
  ) {
    return undefined;
  }
  const sourceFontSize = resolveSpreadsheetCellFontSize(cell);
  const availableWidth = Math.max(
    1,
    (contentWidth - SPREADSHEET_CONTENT_HORIZONTAL_INSET) *
      SPREADSHEET_CONTENT_WIDTH_SAFETY_RATIO,
  );
  const availableHeight = Math.max(
    1,
    contentHeight - SPREADSHEET_CONTENT_VERTICAL_INSET,
  );
  // Excel 单元格常用末尾换行保存编辑状态，但该换行不会占据额外可见行。
  const measuredValue = cell.value.replace(/(?:\r?\n)+$/, '');
  const fits = (fontSize: number) => {
    return (
      estimateSpreadsheetWrappedLineCount(
        measuredValue,
        fontSize,
        availableWidth,
      ) *
        fontSize *
        SPREADSHEET_CONTENT_LINE_HEIGHT <=
      availableHeight
    );
  };
  if (fits(sourceFontSize)) return undefined;

  let lower = 1;
  let upper = sourceFontSize;
  for (let attempt = 0; attempt < 10; attempt += 1) {
    const candidate = (lower + upper) / 2;
    if (fits(candidate)) {
      lower = candidate;
    } else {
      upper = candidate;
    }
  }
  return Math.max(1, Math.floor(lower * 100) / 100);
}

/** 渲染不依赖图标字体的条件格式图标。 */
function SpreadsheetConditionalIcon({ index }: { index: number }) {
  const colors = ['#cf1322', '#d48806', '#389e0d'];
  return (
    <svg
      className="office-file-spreadsheet-conditional-icon"
      viewBox="0 0 12 12"
      aria-hidden="true"
    >
      <circle cx="6" cy="6" r="4.5" fill={colors[index] ?? colors[0]} />
    </svg>
  );
}

/** 渲染电子表格单元格的共享内容层。 */
function SpreadsheetCellRendererComponent({
  cell,
  contentWidth,
  contentHeight,
  clipped,
  contentBounds,
  viewMode = 'source',
  sourceId,
  sheetId,
}: SpreadsheetCellRendererProps) {
  const hyperlinkProps = useOfficeHyperlink<HTMLDivElement>({
    hyperlink: cell.hyperlink,
    source: { type: 'cell', id: sourceId },
  });
  const shrinkToFit = isSpreadsheetShrinkToFitCell(cell);
  const wrapped = Boolean(
    cell.style?.wrapText || (viewMode === 'reading' && !shrinkToFit),
  );
  const fittedFontSize =
    (viewMode === 'source'
      ? resolveBoundedWrapFontSize(cell, contentWidth, contentHeight, clipped)
      : undefined) ?? resolveShrinkToFitFontSize(cell, contentWidth);
  const clippedPosition = clipped
    ? cell.style?.verticalAlign === 'bottom'
      ? { bottom: 0 }
      : cell.style?.verticalAlign === 'middle'
      ? { top: '50%', transform: 'translateY(-50%)' }
      : { top: 0 }
    : undefined;
  const highlightedValue = (
    <OfficeSearchHighlightedText
      text={cell.value}
      target={{
        kind: 'spreadsheet',
        sheetId,
        rowIndex: cell.rowIndex,
        columnIndex: cell.columnIndex,
      }}
    />
  );
  return (
    <>
      {cell.conditionalVisual?.dataBarPercent !== undefined ? (
        <span
          className="office-file-spreadsheet-data-bar"
          aria-hidden="true"
          style={{
            width: `${cell.conditionalVisual.dataBarPercent}%`,
            backgroundColor: cell.conditionalVisual.dataBarColor,
          }}
        />
      ) : null}
      {cell.conditionalVisual?.iconIndex !== undefined ? (
        <SpreadsheetConditionalIcon index={cell.conditionalVisual.iconIndex} />
      ) : null}
      <div
        {...hyperlinkProps}
        className={`office-file-xlsx-sheet-table__cell-content${
          wrapped ? ' office-file-xlsx-sheet-table__cell-content--wrapped' : ''
        }${
          clipped ? ' office-file-xlsx-sheet-table__cell-content--clipped' : ''
        }${
          contentBounds
            ? ' office-file-xlsx-sheet-table__cell-content--bounded'
            : ''
        }`}
        style={{
          maxHeight: contentHeight,
          fontSize: fittedFontSize,
          ...clippedPosition,
          ...(contentBounds
            ? {
                left: contentBounds.left,
                width: contentBounds.width,
                minWidth: 0,
              }
            : undefined),
        }}
      >
        {contentBounds ? (
          <span
            className="office-file-xlsx-sheet-table__cell-text"
            style={{
              left: contentBounds.textLeft,
              width: contentBounds.textWidth,
              justifyContent: contentBounds.textJustify,
            }}
          >
            <span className="office-file-xlsx-sheet-table__cell-text-value">
              {highlightedValue}
            </span>
          </span>
        ) : (
          highlightedValue
        )}
      </div>
      {cell.style?.diagonalBorder ? (
        <SpreadsheetDiagonalBorderLayer border={cell.style.diagonalBorder} />
      ) : null}
      <SpreadsheetAnnotationMarker annotation={cell.annotation} />
      {cell.table?.showFilter ? (
        <span
          className="office-file-spreadsheet-filter-indicator"
          aria-hidden="true"
        />
      ) : null}
    </>
  );
}

export const SpreadsheetCellRenderer = memo(SpreadsheetCellRendererComponent);
