import { createContentStore } from '../content-store';
import type { OfficeContentStore } from '../content-store/types';
import { createSpreadsheetAxisIndex } from './SpreadsheetAxisIndex';
import { SpreadsheetMergeIndex } from './SpreadsheetMergeIndex';
import { SpreadsheetObjectIndex } from './SpreadsheetObjectIndex';
import {
  throwIfSpreadsheetAborted,
  type SpreadsheetSheetLayout,
} from './SpreadsheetSource';
import type {
  SpreadsheetCell,
  SpreadsheetChart,
  SpreadsheetColumnMetric,
  SpreadsheetHyperlinkRange,
  SpreadsheetImage,
  SpreadsheetMerge,
  SpreadsheetRange,
  SpreadsheetRangeData,
  SpreadsheetRowMetric,
} from './types';

/** 稀疏 Sheet Store 使用的固定行分块大小。 */
export const SPREADSHEET_TILE_ROWS = 200;
/** 稀疏 Sheet Store 使用的固定列分块大小。 */
export const SPREADSHEET_TILE_COLUMNS = 50;

/** 描述一个稀疏单元格 tile。 */
export type SpreadsheetTile = {
  /** 工作表的稳定标识。 */
  sheetId: string;
  /** 行 分片。 */
  rowTile: number;
  /** 列 分片。 */
  columnTile: number;
  /** 数据源变更时递增的修订号。 */
  revision: number;
  /** 按显示顺序排列的单元格。 */
  cells: readonly SpreadsheetCell[];
};

/** 描述工作表不随范围重复保存的结构和尺寸。 */
export type SpreadsheetSheetStructure = SpreadsheetSheetLayout & {
  /** 工作表全局声明且不随分片重复保存的合并区域。 */
  merges: readonly SpreadsheetMerge[];
  /** 当前文档或页面包含的图片资源。 */
  images: readonly SpreadsheetImage[];
  /** 按图表对象编号索引的标准图表模型。 */
  charts: readonly SpreadsheetChart[];
  /** 工作表声明的稀疏超链接范围。 */
  hyperlinks?: readonly SpreadsheetHyperlinkRange[];
};

/** Sheet Store 对 Source 暴露的范围读取协议。 */
export interface SpreadsheetSheetStore {
  /** 保存工作表中一个已解析的数据分片。 */
  putTile(tile: SpreadsheetTile): Promise<void>;
  /** 保存工作表的行列结构与合并信息。 */
  putStructure(structure: SpreadsheetSheetStructure): void;
  /** 读取指定工作表的布局信息。 */
  getLayout(): SpreadsheetSheetLayout;
  /** 读取指定工作表范围内的单元格与对象。 */
  getRange(
    range: SpreadsheetRange,
    signal?: AbortSignal,
  ): Promise<SpreadsheetRangeData>;
  /** 保留指定可视范围并回收远离窗口的缓存内容。 */
  retainRange(range: SpreadsheetRange): () => void;
  /** 幂等释放当前对象持有的资源和订阅。 */
  dispose(): Promise<void>;
}

/** 工作表按需分片的访问时间和加载状态。 */
type TileMeta = {
  /** 行 分片。 */
  rowTile: number;
  /** 列 分片。 */
  columnTile: number;
};

function tileKey(sheetId: string, rowTile: number, columnTile: number) {
  return `${sheetId}:${rowTile}:${columnTile}`;
}

function columnLabel(column: number) {
  let value = Math.max(1, Math.floor(column));
  let label = '';
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return label;
}

function normalizeRange(
  range: SpreadsheetRange,
  rowCount: number,
  columnCount: number,
) {
  const startRow = Math.max(1, Math.min(rowCount, Math.floor(range.startRow)));
  const endRow = Math.max(
    startRow,
    Math.min(rowCount, Math.floor(range.endRow)),
  );
  const startColumn = Math.max(
    1,
    Math.min(columnCount, Math.floor(range.startColumn)),
  );
  const endColumn = Math.max(
    startColumn,
    Math.min(columnCount, Math.floor(range.endColumn)),
  );
  return { startRow, endRow, startColumn, endColumn };
}

function enumerateTileKeys(sheetId: string, range: SpreadsheetRange) {
  const startRowTile = Math.floor((range.startRow - 1) / SPREADSHEET_TILE_ROWS);
  const endRowTile = Math.floor((range.endRow - 1) / SPREADSHEET_TILE_ROWS);
  const startColumnTile = Math.floor(
    (range.startColumn - 1) / SPREADSHEET_TILE_COLUMNS,
  );
  const endColumnTile = Math.floor(
    (range.endColumn - 1) / SPREADSHEET_TILE_COLUMNS,
  );
  const keys: string[] = [];
  for (let rowTile = startRowTile; rowTile <= endRowTile; rowTile += 1) {
    for (
      let columnTile = startColumnTile;
      columnTile <= endColumnTile;
      columnTile += 1
    ) {
      keys.push(tileKey(sheetId, rowTile, columnTile));
    }
  }
  return keys;
}

function metricMap<T extends SpreadsheetRowMetric | SpreadsheetColumnMetric>(
  metrics: readonly T[],
) {
  return new Map(metrics.map((metric) => [metric.index, metric]));
}

/** 创建带内存 LRU 和 IndexedDB 冷层的稀疏工作表 Store。 */
export function createSpreadsheetSheetStore(
  sessionId: string,
  sheetId: string,
  initialStructure: SpreadsheetSheetStructure,
): SpreadsheetSheetStore {
  const tiles: OfficeContentStore<TileMeta, SpreadsheetTile> =
    createContentStore({
      sessionId,
      namespace: `spreadsheet-${sheetId}`,
      maxMemoryBytes: 32 * 1024 * 1024,
      estimateSize: (tile) => 256 + tile.cells.length * 384,
    });
  let structure = initialStructure;
  let rowAxis = createSpreadsheetAxisIndex(
    structure.rowCount,
    structure.defaultRowHeight,
    structure.rows,
  );
  let columnAxis = createSpreadsheetAxisIndex(
    structure.columnCount,
    structure.defaultColumnWidth,
    structure.columns,
  );
  let mergeIndex = new SpreadsheetMergeIndex(structure.merges);
  let imageIndex = new SpreadsheetObjectIndex(
    structure.images,
    rowAxis,
    columnAxis,
  );
  let chartIndex = new SpreadsheetObjectIndex(
    structure.charts,
    rowAxis,
    columnAxis,
  );
  let disposed = false;

  const ensureAvailable = () => {
    if (disposed) throw new Error('工作表 Store 已释放');
  };

  return {
    async putTile(tile) {
      ensureAvailable();
      await tiles.put({
        key: tileKey(sheetId, tile.rowTile, tile.columnTile),
        revision: tile.revision,
        meta: {
          rowTile: tile.rowTile,
          columnTile: tile.columnTile,
        },
        value: tile,
        updatedAt: Date.now(),
      });
    },
    putStructure(nextStructure) {
      ensureAvailable();
      if (nextStructure.revision < structure.revision) return;
      structure = nextStructure;
      rowAxis = createSpreadsheetAxisIndex(
        structure.rowCount,
        structure.defaultRowHeight,
        structure.rows,
      );
      columnAxis = createSpreadsheetAxisIndex(
        structure.columnCount,
        structure.defaultColumnWidth,
        structure.columns,
      );
      mergeIndex = new SpreadsheetMergeIndex(structure.merges);
      imageIndex = new SpreadsheetObjectIndex(
        structure.images,
        rowAxis,
        columnAxis,
      );
      chartIndex = new SpreadsheetObjectIndex(
        structure.charts,
        rowAxis,
        columnAxis,
      );
    },
    getLayout() {
      ensureAvailable();
      return {
        revision: structure.revision,
        rowCount: structure.rowCount,
        columnCount: structure.columnCount,
        defaultRowHeight: structure.defaultRowHeight,
        defaultColumnWidth: structure.defaultColumnWidth,
        rows: structure.rows,
        columns: structure.columns,
      };
    },
    async getRange(requestedRange, signal) {
      ensureAvailable();
      throwIfSpreadsheetAborted(signal);
      const normalized = normalizeRange(
        requestedRange,
        structure.rowCount,
        structure.columnCount,
      );
      const range = normalizeRange(
        mergeIndex.expand(normalized),
        structure.rowCount,
        structure.columnCount,
      );
      const records = await Promise.all(
        enumerateTileKeys(sheetId, range).map((key) => tiles.get(key, signal)),
      );
      throwIfSpreadsheetAborted(signal);
      const cells = records.flatMap((record) =>
        (record?.value?.cells ?? [])
          .filter(
            (cell) =>
              cell.rowIndex >= range.startRow &&
              cell.rowIndex <= range.endRow &&
              cell.columnIndex >= range.startColumn &&
              cell.columnIndex <= range.endColumn,
          )
          .map((cell) => ({ ...cell })),
      );
      const cellByRef = new Map(cells.map((cell) => [cell.ref, cell]));
      (structure.hyperlinks ?? []).forEach((rangeLink) => {
        const startRow = Math.max(range.startRow, rangeLink.startRow);
        const endRow = Math.min(range.endRow, rangeLink.endRow);
        const startColumn = Math.max(range.startColumn, rangeLink.startColumn);
        const endColumn = Math.min(range.endColumn, rangeLink.endColumn);
        if (startRow > endRow || startColumn > endColumn) return;
        for (let row = startRow; row <= endRow; row += 1) {
          for (let column = startColumn; column <= endColumn; column += 1) {
            const ref = `${columnLabel(column)}${row}`;
            const existing = cellByRef.get(ref);
            if (existing) {
              existing.hyperlink = rangeLink.hyperlink;
            } else {
              const linkedCell: SpreadsheetCell = {
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
      const rowsByIndex = metricMap(structure.rows);
      const columnsByIndex = metricMap(structure.columns);
      const rows = Array.from(
        { length: range.endRow - range.startRow + 1 },
        (_, offset): SpreadsheetRowMetric => {
          const index = range.startRow + offset;
          return (
            rowsByIndex.get(index) ?? {
              index,
              height: structure.defaultRowHeight,
              hidden: false,
            }
          );
        },
      );
      const columns = Array.from(
        { length: range.endColumn - range.startColumn + 1 },
        (_, offset): SpreadsheetColumnMetric => {
          const index = range.startColumn + offset;
          return (
            columnsByIndex.get(index) ?? {
              index,
              width: structure.defaultColumnWidth,
              hidden: false,
            }
          );
        },
      );
      return {
        revision: Math.max(
          structure.revision,
          ...records.map((record) => record?.revision ?? 0),
        ),
        range,
        cells,
        rows,
        columns,
        merges: mergeIndex.query(range),
        images: imageIndex.query(range),
        charts: chartIndex.query(range),
      };
    },
    retainRange(range) {
      ensureAvailable();
      const normalized = normalizeRange(
        range,
        structure.rowCount,
        structure.columnCount,
      );
      return tiles.pin(enumerateTileKeys(sheetId, normalized));
    },
    async dispose() {
      if (disposed) return;
      disposed = true;
      await tiles.dispose();
    },
  };
}
