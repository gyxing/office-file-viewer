// SpreadsheetCellRenderer 统一渲染普通表格与窗口表格的单元格内容和对角边框。
import React, { memo } from 'react';
import type {
  SpreadsheetCell,
  SpreadsheetDiagonalBorder,
} from '../../services/spreadsheet/types';
import type { SpreadsheetCellContentBounds } from './spreadsheetCellOverflow';

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

/** 固定边界内容左右内边距与折叠边框占用的近似宽度。 */
const BOUNDED_CONTENT_HORIZONTAL_INSET = 5;
/** 固定边界内容为避免压住底边保留的垂直空间。 */
const BOUNDED_CONTENT_VERTICAL_INSET = 1;
/** 与固定边界内容的 CSS 行高保持一致。 */
const BOUNDED_CONTENT_LINE_HEIGHT = 1.2;

/** 粗略估算单个字符宽度，避免为大量单元格创建 DOM 测量节点。 */
function estimateSpreadsheetCharacterWidth(
  character: string,
  fontSize: number,
) {
  if (/\s/.test(character)) return fontSize * 0.25;
  if (
    /[\u2e80-\u9fff\uf900-\ufaff\uff01-\uff60\uffe0-\uffe6]/.test(character)
  ) {
    return fontSize;
  }
  if (/[MW]/.test(character)) return fontSize * 0.9;
  if (/[CDGOQ]/.test(character)) return fontSize * 0.75;
  if (/[PFLT]/.test(character)) return fontSize * 0.57;
  if (/[IJ]/.test(character)) return fontSize * 0.36;
  if (/[A-Z]/.test(character)) return fontSize * 0.66;
  if (/[a-z0-9]/.test(character)) return fontSize * 0.5;
  return fontSize * 0.5;
}

/** 粗略估算单行文本宽度，避免为大量单元格创建 DOM 测量节点。 */
function estimateSpreadsheetLineWidth(value: string, fontSize: number) {
  return [...value].reduce(
    (width, character) =>
      width + estimateSpreadsheetCharacterWidth(character, fontSize),
    0,
  );
}

/** 按 Excel 的空白优先、超长单词按字符拆分规则估算换行数量。 */
function estimateWrappedLineCount(
  value: string,
  fontSize: number,
  availableWidth: number,
) {
  let lineCount = 0;
  value.split(/\r?\n/).forEach((sourceLine) => {
    const tokens = sourceLine.match(/\s+|\S+/g) ?? [''];
    let currentWidth = 0;
    lineCount += 1;
    tokens.forEach((token) => {
      const tokenWidth = estimateSpreadsheetLineWidth(token, fontSize);
      if (currentWidth > 0 && currentWidth + tokenWidth > availableWidth) {
        lineCount += 1;
        currentWidth = 0;
      }
      if (tokenWidth <= availableWidth) {
        currentWidth += tokenWidth;
        return;
      }
      [...token].forEach((character) => {
        const characterWidth = estimateSpreadsheetCharacterWidth(
          character,
          fontSize,
        );
        if (
          currentWidth > 0 &&
          currentWidth + characterWidth > availableWidth
        ) {
          lineCount += 1;
          currentWidth = 0;
        }
        currentWidth += characterWidth;
      });
    });
  });
  return lineCount;
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
  const fallbackFontSize =
    style.bold || style.color?.toLowerCase() === '#ff0000' ? 14 : 13;
  const fontSize = style.fontSize ?? fallbackFontSize;
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
  const fallbackFontSize =
    style.bold || style.color?.toLowerCase() === '#ff0000' ? 14 : 13;
  const sourceFontSize = style.fontSize ?? fallbackFontSize;
  const availableWidth = Math.max(
    1,
    contentWidth - BOUNDED_CONTENT_HORIZONTAL_INSET,
  );
  const availableHeight = Math.max(
    1,
    contentHeight - BOUNDED_CONTENT_VERTICAL_INSET,
  );
  // Excel 单元格常用末尾换行保存编辑状态，但该换行不会占据额外可见行。
  const measuredValue = cell.value.replace(/(?:\r?\n)+$/, '');
  const fits = (fontSize: number) => {
    return (
      estimateWrappedLineCount(measuredValue, fontSize, availableWidth) *
        fontSize *
        BOUNDED_CONTENT_LINE_HEIGHT <=
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

/** 渲染电子表格单元格的共享内容层。 */
function SpreadsheetCellRendererComponent({
  cell,
  contentWidth,
  contentHeight,
  clipped,
  contentBounds,
}: SpreadsheetCellRendererProps) {
  const fittedFontSize =
    resolveBoundedWrapFontSize(cell, contentWidth, contentHeight, clipped) ??
    resolveShrinkToFitFontSize(cell, contentWidth);
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
          cell.style?.wrapText
            ? ' office-file-xlsx-sheet-table__cell-content--wrapped'
            : ''
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
              {cell.value}
            </span>
          </span>
        ) : (
          cell.value
        )}
      </div>
      {cell.style?.diagonalBorder ? (
        <SpreadsheetDiagonalBorderLayer border={cell.style.diagonalBorder} />
      ) : null}
    </>
  );
}

export const SpreadsheetCellRenderer = memo(SpreadsheetCellRendererComponent);
