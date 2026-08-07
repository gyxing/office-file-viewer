import type { SpreadsheetAxisIndex } from '../../services/spreadsheet/SpreadsheetAxisIndex';
import type {
  SpreadsheetCell,
  SpreadsheetRangeData,
  SpreadsheetRowMetric,
} from '../../services/spreadsheet/types';

/** 单元格左右内边距和折叠边框占用的近似宽度。 */
export const SPREADSHEET_CONTENT_HORIZONTAL_INSET = 5;
/** 单元格为避免字形压住底边保留的垂直空间。 */
export const SPREADSHEET_CONTENT_VERTICAL_INSET = 1;
/** 与单元格换行内容的 CSS 行高保持一致。 */
export const SPREADSHEET_CONTENT_LINE_HEIGHT = 1.2;
/** 为浏览器字形宽度与静态估算之间的误差预留空间。 */
export const SPREADSHEET_CONTENT_WIDTH_SAFETY_RATIO = 0.92;

/** 按单元格样式返回渲染层使用的兼容字号。 */
export function resolveSpreadsheetCellFontSize(cell: SpreadsheetCell) {
  const style = cell.style;
  const fallbackFontSize =
    style?.bold || style?.color?.toLowerCase() === '#ff0000' ? 14 : 13;
  return style?.fontSize ?? fallbackFontSize;
}

/** 判断源样式是否实际采用单行缩小字体；自动换行存在时按换行优先。 */
export function isSpreadsheetShrinkToFitCell(cell: SpreadsheetCell) {
  return Boolean(cell.style?.shrinkToFit && !cell.style.wrapText);
}

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
  return fontSize * 0.5;
}

/** 粗略估算单行文本宽度。 */
export function estimateSpreadsheetLineWidth(value: string, fontSize: number) {
  return [...value].reduce(
    (width, character) =>
      width + estimateSpreadsheetCharacterWidth(character, fontSize),
    0,
  );
}

/** 按空白优先、超长单词按字符拆分的规则估算换行数量。 */
export function estimateSpreadsheetWrappedLineCount(
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
  return Math.max(1, lineCount);
}

/** 估算阅读模式完整显示文本所需的单元格边框盒高度。 */
export function calculateSpreadsheetReadingTextHeight(
  cell: SpreadsheetCell,
  contentWidth: number,
) {
  const fontSize = resolveSpreadsheetCellFontSize(cell);
  const availableWidth = Math.max(
    1,
    (contentWidth - SPREADSHEET_CONTENT_HORIZONTAL_INSET) *
      SPREADSHEET_CONTENT_WIDTH_SAFETY_RATIO,
  );
  // Office 常在末尾保留编辑态换行，但该换行不应单独撑出空白行。
  const measuredValue = cell.value.replace(/(?:\r?\n)+$/, '');
  const lineCount = estimateSpreadsheetWrappedLineCount(
    measuredValue,
    fontSize,
    availableWidth,
  );
  return Math.ceil(
    lineCount * fontSize * SPREADSHEET_CONTENT_LINE_HEIGHT +
      SPREADSHEET_CONTENT_VERTICAL_INSET,
  );
}

/** 计算当前加载范围在阅读模式下新增或变大的稀疏行高。 */
export function buildSpreadsheetReadingRowHeightUpdates(
  data: SpreadsheetRangeData,
  rowAxis: SpreadsheetAxisIndex,
  columnAxis: SpreadsheetAxisIndex,
) {
  const updates = new Map<number, number>();
  const hiddenRows = new Set(
    data.rows.filter((row) => row.hidden).map((row) => row.index),
  );
  const mergeByAnchor = new Map(
    data.merges.map((merge) => [
      `${merge.startRow}:${merge.startColumn}`,
      merge,
    ]),
  );

  data.cells.forEach((cell) => {
    if (
      !cell.value ||
      cell.hiddenByMerge ||
      isSpreadsheetShrinkToFitCell(cell)
    ) {
      return;
    }
    const merge = mergeByAnchor.get(`${cell.rowIndex}:${cell.columnIndex}`);
    const endRow = merge?.endRow ?? cell.rowIndex;
    const endColumn = merge?.endColumn ?? cell.columnIndex;
    const targetRow = (() => {
      for (let row = endRow; row >= cell.rowIndex; row -= 1) {
        if (!hiddenRows.has(row) && rowAxis.sizeAt(row) > 0) return row;
      }
      return undefined;
    })();
    if (!targetRow) return;

    const requiredHeight = calculateSpreadsheetReadingTextHeight(
      cell,
      columnAxis.rangeSize(cell.columnIndex, endColumn),
    );
    const currentRangeHeight = rowAxis.rangeSize(cell.rowIndex, endRow);
    if (requiredHeight <= currentRangeHeight) return;
    const targetHeight =
      rowAxis.sizeAt(targetRow) + requiredHeight - currentRangeHeight;
    updates.set(targetRow, Math.max(updates.get(targetRow) ?? 0, targetHeight));
  });
  return updates;
}

/** 将源行尺寸与阅读模式缓存合并为唯一、升序的轴尺寸集合。 */
export function mergeSpreadsheetReadingRowMetrics(
  sourceRows: readonly SpreadsheetRowMetric[],
  readingRowHeights: ReadonlyMap<number, number>,
) {
  const metrics = new Map(sourceRows.map((row) => [row.index, row]));
  readingRowHeights.forEach((height, index) => {
    const source = metrics.get(index);
    if (source?.hidden) return;
    metrics.set(index, {
      index,
      height: Math.max(source?.height ?? 0, height),
      customHeight: false,
      hidden: false,
    });
  });
  return [...metrics.values()].sort((left, right) => left.index - right.index);
}

/** 将源行轴中的浮动对象纵向范围映射到阅读模式行轴。 */
export function remapSpreadsheetVerticalRange(
  y: number,
  height: number,
  sourceAxis: SpreadsheetAxisIndex,
  displayAxis: SpreadsheetAxisIndex,
) {
  const remapOffset = (offset: number) => {
    const rowIndex = sourceAxis.findIndexAtOffset(offset);
    const sourceStart = sourceAxis.offsetAt(rowIndex);
    const sourceSize = sourceAxis.sizeAt(rowIndex);
    const ratio =
      sourceSize > 0
        ? Math.max(0, Math.min(1, (offset - sourceStart) / sourceSize))
        : 0;
    return (
      displayAxis.offsetAt(rowIndex) + displayAxis.sizeAt(rowIndex) * ratio
    );
  };
  const start = remapOffset(y);
  const end = remapOffset(y + height);
  return { y: start, height: Math.max(1, end - start) };
}
