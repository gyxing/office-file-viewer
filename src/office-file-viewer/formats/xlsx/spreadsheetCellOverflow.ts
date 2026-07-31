// spreadsheetCellOverflow 统一计算普通表格与虚拟表格的单元格文本溢出走廊。
import type { SpreadsheetCell } from '../../services/spreadsheet/types';

/** 描述文本可见走廊及其在源单元格中的对齐基准。 */
export type SpreadsheetCellContentBounds = {
  /** 可见走廊相对源单元格内容起点的偏移。 */
  left: number;
  /** 可见走廊宽度。 */
  width: number;
  /** 文本对齐基准相对可见走廊起点的偏移。 */
  textLeft: number;
  /** 文本仍按源单元格内容宽度进行对齐。 */
  textWidth: number;
  /** 在源单元格宽度内保持 Excel 的文本锚定方向。 */
  textJustify: 'start' | 'center' | 'end';
};

/** 描述一行内参与文本溢出计算的单元格。 */
export type SpreadsheetRowCellLayout = {
  /** 用于稳定识别或缓存当前项目的键。 */
  key: string;
  /** 当前处理的单元格。 */
  cell: SpreadsheetCell;
  /** 当前单元格相对合并区域起始列的偏移。 */
  columnOffset: number;
  /** 单元格横向跨越的列数。 */
  columnSpan?: number;
  /** 内容是否需要限制在指定边界内。 */
  clipped?: boolean;
};

// 与单元格左右各 2px 的样式保持一致，避免走廊侵入相邻单元格内边距。
/** 计算单元格文字溢出时预留的水平内边距，单位为标准化渲染像素。 */
const CELL_HORIZONTAL_PADDING = 4;

/** 判断单元格是否会阻挡相邻文本继续溢出。 */
export function isSpreadsheetCellOccupied(cell: SpreadsheetCell | undefined) {
  if (!cell) return false;
  return Boolean(
    cell.value !== '' ||
      cell.formula ||
      (cell.rawValue !== undefined && cell.rawValue !== ''),
  );
}

/** 判断常规对齐下的值是否属于可向右溢出的文本。 */
function isTextCell(cell: SpreadsheetCell) {
  return (
    cell.type === 's' ||
    cell.type === 'str' ||
    cell.type === 'inlineStr' ||
    cell.type === 'string' ||
    cell.type === 'e' ||
    cell.type === 'error'
  );
}

/** 确定文本相对源单元格允许延伸的方向。 */
function resolveOverflowDirection(cell: SpreadsheetCell) {
  const alignment = cell.style?.horizontalAlign;
  if (alignment === 'left') return 'right';
  if (alignment === 'right') return 'left';
  if (alignment === 'center') return 'both';
  if (alignment === 'justify') return 'none';
  return isTextCell(cell) ? 'right' : 'none';
}

/**
 * 为一行单元格批量生成文本可见走廊。
 *
 * 最近阻挡列通过前后各一次线性扫描得到，避免在大文件窗口中逐格向两侧查找。
 */
export function buildSpreadsheetRowContentBounds(
  cells: readonly SpreadsheetRowCellLayout[],
  occupiedColumns: readonly boolean[],
  columnWidths: readonly number[],
) {
  const bounds = new Map<string, SpreadsheetCellContentBounds>();
  const columnCount = Math.min(occupiedColumns.length, columnWidths.length);
  if (!columnCount) return bounds;

  const prefixWidths = new Array<number>(columnCount + 1).fill(0);
  for (let index = 0; index < columnCount; index += 1) {
    prefixWidths[index + 1] =
      prefixWidths[index] + Math.max(0, columnWidths[index] ?? 0);
  }

  const previousOccupied = new Array<number>(columnCount).fill(-1);
  let previous = -1;
  for (let index = 0; index < columnCount; index += 1) {
    previousOccupied[index] = previous;
    if (occupiedColumns[index]) previous = index;
  }

  const nextOccupied = new Array<number>(columnCount).fill(columnCount);
  let next = columnCount;
  for (let index = columnCount - 1; index >= 0; index -= 1) {
    nextOccupied[index] = next;
    if (occupiedColumns[index]) next = index;
  }

  cells.forEach(({ key, cell, columnOffset, columnSpan = 1, clipped }) => {
    if (!cell.value || clipped || columnOffset >= columnCount) return;
    const start = Math.max(0, columnOffset);
    const end = Math.min(columnCount - 1, start + columnSpan - 1);
    const cellWidth = Math.max(
      0,
      prefixWidths[end + 1] - prefixWidths[start] - CELL_HORIZONTAL_PADDING,
    );
    const leftBlocker = previousOccupied[start];
    const rightBlocker = nextOccupied[end];
    const leftAvailable = prefixWidths[start] - prefixWidths[leftBlocker + 1];
    const rightAvailable = prefixWidths[rightBlocker] - prefixWidths[end + 1];
    const direction = resolveOverflowDirection(cell);

    if (direction === 'right') {
      // 右侧没有阻挡项时保留自然溢出，允许文本继续进入工作表的空白补齐区域。
      if (rightBlocker >= columnCount) return;
      bounds.set(key, {
        left: 0,
        width: cellWidth + rightAvailable,
        textLeft: 0,
        textWidth: cellWidth,
        textJustify: 'start',
      });
      return;
    }

    if (direction === 'left') {
      bounds.set(key, {
        left: -leftAvailable,
        width: cellWidth + leftAvailable,
        textLeft: leftAvailable,
        textWidth: cellWidth,
        textJustify: 'end',
      });
      return;
    }

    if (direction === 'both') {
      bounds.set(key, {
        left: -leftAvailable,
        width: cellWidth + leftAvailable + rightAvailable,
        textLeft: leftAvailable,
        textWidth: cellWidth,
        textJustify: 'center',
      });
      return;
    }

    // 数字、布尔值等常规对齐内容不会像文本一样跨越相邻空单元格。
    bounds.set(key, {
      left: 0,
      width: cellWidth,
      textLeft: 0,
      textWidth: cellWidth,
      textJustify:
        cell.style?.horizontalAlign === 'right'
          ? 'end'
          : cell.style?.horizontalAlign === 'center'
          ? 'center'
          : 'start',
    });
  });

  return bounds;
}
