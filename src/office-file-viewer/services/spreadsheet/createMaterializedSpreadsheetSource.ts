import { SpreadsheetMergeIndex } from './SpreadsheetMergeIndex';
import {
  createSpreadsheetPerformanceProfile,
  type SpreadsheetPerformanceProfile,
} from './spreadsheetPerformance';
import { SpreadsheetSearchProvider } from './SpreadsheetSearchProvider';
import type {
  SpreadsheetSheetDescriptor,
  SpreadsheetSheetLayout,
  SpreadsheetSource,
  SpreadsheetSourceSnapshot,
} from './SpreadsheetSource';
import { throwIfSpreadsheetAborted } from './SpreadsheetSource';
import type {
  SpreadsheetRange,
  SpreadsheetRangeData,
  SpreadsheetSheet,
  SpreadsheetWorkbook,
} from './types';

function rangeIntersectsObject(
  range: SpreadsheetRange,
  object: {
    from: { row: number; column: number };
    to: { row: number; column: number };
  },
) {
  return !(
    object.to.row < range.startRow ||
    object.from.row > range.endRow ||
    object.to.column < range.startColumn ||
    object.from.column > range.endColumn
  );
}

function cellReference(row: number, column: number) {
  let current = column;
  let label = '';
  while (current > 0) {
    current -= 1;
    label = String.fromCharCode(65 + (current % 26)) + label;
    current = Math.floor(current / 26);
  }
  return `${label}${row}`;
}

/** 为完整工作簿创建零拷贝 Source 适配器。 */
export function createMaterializedSpreadsheetSource(
  workbook: SpreadsheetWorkbook,
): SpreadsheetSource {
  const descriptors: SpreadsheetSheetDescriptor[] = workbook.sheets.map(
    (sheet) => ({
      id: sheet.id,
      name: sheet.name,
      path: sheet.path,
      kind: sheet.kind ?? 'worksheet',
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      revision: 1,
      status: 'ready',
    }),
  );
  const profiles = new Map<string, SpreadsheetPerformanceProfile>(
    workbook.sheets.map((sheet) => [
      sheet.id,
      createSpreadsheetPerformanceProfile({
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
      }),
    ]),
  );
  const snapshot: SpreadsheetSourceSnapshot = {
    revision: 1,
    sheets: descriptors,
    definedNames: workbook.definedNames,
  };
  const sheetById = new Map(workbook.sheets.map((sheet) => [sheet.id, sheet]));

  const getSheet = (sheetId: string) => {
    const sheet = sheetById.get(sheetId);
    if (!sheet) throw new RangeError(`工作表不存在：${sheetId}`);
    return sheet;
  };
  const getProfile = (sheetId: string) => {
    const profile = profiles.get(sheetId);
    if (!profile) throw new RangeError(`工作表不存在：${sheetId}`);
    return profile;
  };
  const getLayout = (sheet: SpreadsheetSheet): SpreadsheetSheetLayout => ({
    revision: 1,
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
    defaultRowHeight: sheet.defaultRowHeight ?? 20,
    defaultColumnWidth: sheet.defaultColumnWidth ?? 64,
    rows: sheet.rows.map((row) => ({
      index: row.index,
      height: row.height,
      customHeight: row.customHeight,
      hidden: Boolean(row.hidden),
    })),
    columns: sheet.columns.map((column) => ({
      index: column.index,
      width: column.width,
      hidden: Boolean(column.hidden),
    })),
  });

  let searchProvider: SpreadsheetSearchProvider;
  const source: SpreadsheetSource = {
    get searchProvider() {
      return searchProvider;
    },
    getSnapshot: () => snapshot,
    subscribe: () => () => undefined,
    getProfile,
    getSheetLayout: (sheetId) => getLayout(getSheet(sheetId)),
    async ensureSheet(sheetId, signal) {
      throwIfSpreadsheetAborted(signal);
      getSheet(sheetId);
    },
    async getMaterializedSheet(sheetId, signal) {
      throwIfSpreadsheetAborted(signal);
      return getProfile(sheetId).gridMode === 'table'
        ? getSheet(sheetId)
        : undefined;
    },
    async getRange(sheetId, requestedRange, signal) {
      throwIfSpreadsheetAborted(signal);
      const sheet = getSheet(sheetId);
      const mergeIndex = new SpreadsheetMergeIndex(sheet.merges);
      const range = mergeIndex.expand({
        startRow: Math.max(1, requestedRange.startRow),
        endRow: Math.min(sheet.rowCount, requestedRange.endRow),
        startColumn: Math.max(1, requestedRange.startColumn),
        endColumn: Math.min(sheet.columnCount, requestedRange.endColumn),
      });
      const rows = sheet.rows.filter(
        (row) => row.index >= range.startRow && row.index <= range.endRow,
      );
      const cells = rows.flatMap((row) =>
        row.cells
          .filter(
            (cell) =>
              cell.columnIndex >= range.startColumn &&
              cell.columnIndex <= range.endColumn,
          )
          .map((cell) => ({ ...cell })),
      );
      const cellByRef = new Map(cells.map((cell) => [cell.ref, cell]));
      (sheet.hyperlinks ?? []).forEach((rangeLink) => {
        const startRow = Math.max(range.startRow, rangeLink.startRow);
        const endRow = Math.min(range.endRow, rangeLink.endRow);
        const startColumn = Math.max(range.startColumn, rangeLink.startColumn);
        const endColumn = Math.min(range.endColumn, rangeLink.endColumn);
        for (let row = startRow; row <= endRow; row += 1) {
          for (let column = startColumn; column <= endColumn; column += 1) {
            const ref = cellReference(row, column);
            const cell = cellByRef.get(ref);
            if (cell) {
              cell.hyperlink = rangeLink.hyperlink;
            } else {
              const linkedCell = {
                ref,
                rowIndex: row,
                columnIndex: column,
                value: '',
                hyperlink: rangeLink.hyperlink,
              };
              cells.push(linkedCell);
              cellByRef.set(ref, linkedCell);
            }
          }
        }
      });
      const data: SpreadsheetRangeData = {
        revision: 1,
        range,
        cells,
        rows: rows.map((row) => ({
          index: row.index,
          height: row.height,
          customHeight: row.customHeight,
          hidden: Boolean(row.hidden),
        })),
        columns: sheet.columns
          .filter(
            (column) =>
              column.index >= range.startColumn &&
              column.index <= range.endColumn,
          )
          .map((column) => ({
            index: column.index,
            width: column.width,
            hidden: Boolean(column.hidden),
          })),
        merges: mergeIndex.query(range),
        images: sheet.images.filter((image) =>
          rangeIntersectsObject(range, image),
        ),
        charts: sheet.charts.filter((chart) =>
          rangeIntersectsObject(range, chart),
        ),
      };
      return data;
    },
    retainRange: () => () => undefined,
    retrySheet: () => undefined,
    dispose: () => Promise.resolve(),
  };
  searchProvider = new SpreadsheetSearchProvider(source);
  return source;
}
