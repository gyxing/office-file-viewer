import {
  OfficeSearchBatchWriter,
  throwIfOfficeSearchAborted,
} from '../search/OfficeSearchProvider';
import {
  createSearchPreviewText,
  findSearchMatches,
  normalizeSearchText,
} from '../search/normalizeSearchText';
import type {
  OfficeSearchProgressEmitter,
  OfficeSearchProvider,
  OfficeSearchQuery,
  OfficeSearchResult,
} from '../search/types';
import type { SpreadsheetSource } from './SpreadsheetSource';

/** 单次范围读取的最大行数，避免大型工作表被物化为稠密矩阵。 */
const SPREADSHEET_SEARCH_ROW_CHUNK_SIZE = 256;

/** 按工作簿、工作表、行和列顺序扫描格式化显示值。 */
export class SpreadsheetSearchProvider implements OfficeSearchProvider {
  readonly kind = 'spreadsheet' as const;

  constructor(private readonly source: SpreadsheetSource) {}

  async search(
    query: OfficeSearchQuery,
    emit: OfficeSearchProgressEmitter,
    signal: AbortSignal,
  ) {
    throwIfOfficeSearchAborted(signal);
    const initialSheets = this.source
      .getSnapshot()
      .sheets.filter((sheet) => sheet.kind === 'worksheet');
    const initialTotal = initialSheets.reduce(
      (total, sheet) => total + Math.max(0, sheet.rowCount),
      0,
    );
    const writer = new OfficeSearchBatchWriter(emit, signal, initialTotal);
    let knownTotal = initialTotal;
    if (!normalizeSearchText(query.text, query.matchCase).text) {
      writer.complete();
      return;
    }

    const visitedCells = new Set<string>();
    for (const initialSheet of initialSheets) {
      throwIfOfficeSearchAborted(signal);
      await this.source.ensureSheet(initialSheet.id, signal);
      const currentSheet =
        this.source
          .getSnapshot()
          .sheets.find((sheet) => sheet.id === initialSheet.id) ?? initialSheet;
      const layout = this.source.getSheetLayout(currentSheet.id);
      const rowCount = Math.max(currentSheet.rowCount, layout.rowCount);
      const columnCount = Math.max(
        currentSheet.columnCount,
        layout.columnCount,
      );
      knownTotal += Math.max(0, rowCount) - Math.max(0, initialSheet.rowCount);
      writer.setTotal(knownTotal);
      if (rowCount <= 0 || columnCount <= 0) continue;

      for (
        let startRow = 1;
        startRow <= rowCount;
        startRow += SPREADSHEET_SEARCH_ROW_CHUNK_SIZE
      ) {
        throwIfOfficeSearchAborted(signal);
        const endRow = Math.min(
          rowCount,
          startRow + SPREADSHEET_SEARCH_ROW_CHUNK_SIZE - 1,
        );
        const rangeData = await this.source.getRange(
          currentSheet.id,
          { startRow, endRow, startColumn: 1, endColumn: columnCount },
          signal,
        );
        const hiddenRows = new Set(
          rangeData.rows.filter((row) => row.hidden).map((row) => row.index),
        );
        const hiddenColumns = new Set(
          rangeData.columns
            .filter((column) => column.hidden)
            .map((column) => column.index),
        );
        const cells = [...rangeData.cells].sort(
          (left, right) =>
            left.rowIndex - right.rowIndex ||
            left.columnIndex - right.columnIndex,
        );

        for (const cell of cells) {
          const cellKey = `${currentSheet.id}:${cell.rowIndex}:${cell.columnIndex}`;
          if (
            visitedCells.has(cellKey) ||
            cell.hiddenByMerge ||
            hiddenRows.has(cell.rowIndex) ||
            hiddenColumns.has(cell.columnIndex)
          ) {
            continue;
          }
          visitedCells.add(cellKey);
          // Excel 的导航目标只精确到单元格，同一单元格内的重复命中合并为一条结果。
          const [match] = findSearchMatches(cell.value, query);
          const items: OfficeSearchResult[] = match
            ? [
                {
                  id: `spreadsheet:${cellKey}`,
                  matchText: cell.value.slice(
                    match.startOffset,
                    match.endOffset,
                  ),
                  previewText: createSearchPreviewText(
                    cell.value,
                    match.startOffset,
                    match.endOffset,
                  ),
                  target: {
                    kind: 'spreadsheet',
                    sheetId: currentSheet.id,
                    rowIndex: cell.rowIndex,
                    columnIndex: cell.columnIndex,
                  },
                },
              ]
            : [];
          await writer.append(items, 0);
        }
        await writer.append([], endRow - startRow + 1);
      }
    }
    writer.complete();
  }
}
