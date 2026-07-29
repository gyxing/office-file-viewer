import { SpreadsheetMergeIndex } from './SpreadsheetMergeIndex';
import {
  createSpreadsheetPerformanceProfile,
  type SpreadsheetPerformanceProfile,
} from './spreadsheetPerformance';
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
      hidden: Boolean(row.hidden),
    })),
    columns: sheet.columns.map((column) => ({
      index: column.index,
      width: column.width,
      hidden: Boolean(column.hidden),
    })),
  });

  return {
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
      const data: SpreadsheetRangeData = {
        revision: 1,
        range,
        cells: rows.flatMap((row) =>
          row.cells.filter(
            (cell) =>
              cell.columnIndex >= range.startColumn &&
              cell.columnIndex <= range.endColumn,
          ),
        ),
        rows: rows.map((row) => ({
          index: row.index,
          height: row.height,
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
}
