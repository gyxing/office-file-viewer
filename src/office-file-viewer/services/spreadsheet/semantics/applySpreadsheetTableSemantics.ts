import type { SpreadsheetCell, SpreadsheetRange } from '../types';
import type { SpreadsheetAutoFilter, SpreadsheetTable } from './types';

/** Excel 内置 Table 样式使用的常见主题色。 */
const TABLE_ACCENT_COLORS = [
  '#4472c4',
  '#ed7d31',
  '#a5a5a5',
  '#ffc000',
  '#5b9bd5',
  '#70ad47',
];

/** 将一基列号转换为 A1 列标签。 */
function columnLabel(index: number) {
  let value = Math.max(1, Math.trunc(index));
  let label = '';
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return label;
}

/** 从内置样式名称稳定选择接近 Excel 主题的强调色。 */
function resolveTableAccent(table: SpreadsheetTable) {
  const styleNumber = Number(table.styleName?.match(/(\d+)$/)?.[1] ?? 1);
  return TABLE_ACCENT_COLORS[(Math.max(1, styleNumber) - 1) % 6];
}

/** 计算表格数据条纹使用的浅色背景。 */
function tintHexColor(color: string, amount = 0.86) {
  const source = color.replace('#', '');
  const parts = [0, 2, 4].map((offset) =>
    Number.parseInt(source.slice(offset, offset + 2), 16),
  );
  return `#${parts
    .map((part) =>
      Math.round(part + (255 - part) * amount)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`;
}

/** 确保当前可见范围内的 Table 空单元格也具有可渲染模型。 */
function ensureCell(
  cells: SpreadsheetCell[],
  cellByPosition: Map<string, SpreadsheetCell>,
  row: number,
  column: number,
) {
  const key = `${row}:${column}`;
  let cell = cellByPosition.get(key);
  if (!cell) {
    cell = {
      ref: `${columnLabel(column)}${row}`,
      rowIndex: row,
      columnIndex: column,
      value: '',
    };
    cells.push(cell);
    cellByPosition.set(key, cell);
  }
  return cell;
}

/** 把 Table、条纹、汇总行和 AutoFilter 静态语义应用到当前范围。 */
export function applySpreadsheetTableSemantics(
  cells: SpreadsheetCell[],
  visibleRange: SpreadsheetRange,
  tables: readonly SpreadsheetTable[],
  autoFilter?: SpreadsheetAutoFilter,
) {
  const cellByPosition = new Map(
    cells.map((cell) => [`${cell.rowIndex}:${cell.columnIndex}`, cell]),
  );
  tables.forEach((table) => {
    const startRow = Math.max(visibleRange.startRow, table.range.startRow);
    const endRow = Math.min(visibleRange.endRow, table.range.endRow);
    const startColumn = Math.max(
      visibleRange.startColumn,
      table.range.startColumn,
    );
    const endColumn = Math.min(visibleRange.endColumn, table.range.endColumn);
    if (startRow > endRow || startColumn > endColumn) return;
    const accent = resolveTableAccent(table);
    const stripe = tintHexColor(accent);
    const headerRow = table.range.startRow;
    const totalsRow = table.totalsRow ? table.range.endRow : undefined;
    for (let row = startRow; row <= endRow; row += 1) {
      for (let column = startColumn; column <= endColumn; column += 1) {
        const cell = ensureCell(cells, cellByPosition, row, column);
        const role =
          table.headerRow && row === headerRow
            ? 'header'
            : totalsRow === row
            ? 'totals'
            : 'body';
        cell.table = {
          tableId: table.id,
          role,
          showFilter: role === 'header' && Boolean(table.autoFilterRef),
        };
        if (role === 'header') {
          cell.style = {
            ...cell.style,
            bold: true,
            color: cell.style?.color ?? '#fff',
            backgroundColor: cell.style?.backgroundColor ?? accent,
          };
        } else if (role === 'totals') {
          cell.style = {
            ...cell.style,
            bold: true,
            borderTop: cell.style?.borderTop ?? `2px double ${accent}`,
          };
        } else if (
          table.showRowStripes &&
          (row - headerRow - 1) % 2 === 1 &&
          !cell.style?.backgroundColor
        ) {
          cell.style = { ...cell.style, backgroundColor: stripe };
        }
      }
    }
  });

  const filterRange = autoFilter?.range;
  if (!filterRange) return;
  const row = filterRange.startRow;
  if (row < visibleRange.startRow || row > visibleRange.endRow) return;
  const startColumn = Math.max(
    visibleRange.startColumn,
    filterRange.startColumn,
  );
  const endColumn = Math.min(visibleRange.endColumn, filterRange.endColumn);
  for (let column = startColumn; column <= endColumn; column += 1) {
    const cell = ensureCell(cells, cellByPosition, row, column);
    cell.table ??= { role: 'header', showFilter: true };
  }
}
