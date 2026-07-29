import {
  createSpreadsheetPerformanceProfile,
  upgradeSpreadsheetPerformanceProfile,
  type SpreadsheetPerformanceProfile,
} from '../spreadsheet/spreadsheetPerformance';
import {
  createSpreadsheetSheetStore,
  type SpreadsheetSheetStore,
} from '../spreadsheet/SpreadsheetSheetStore';
import {
  throwIfSpreadsheetAborted,
  waitForSpreadsheetResult,
  type SpreadsheetSheetDescriptor,
  type SpreadsheetSource,
  type SpreadsheetSourceSnapshot,
} from '../spreadsheet/SpreadsheetSource';
import type {
  SpreadsheetCell,
  SpreadsheetRange,
  SpreadsheetSheet,
} from '../spreadsheet/types';
import { loadXlsxDrawingObjects } from './loadXlsxDrawingObjects';
import { columnIndexToLabel } from './parseXlsx';
import { parseXlsxSheetStream } from './parseXlsxSheetStream';
import {
  profileXlsxArchive,
  readXlsxStructure,
  type ProfiledXlsxArchive,
} from './readXlsxStructure';
import type {
  XlsxPackageContext,
  XlsxSheetDescriptor,
} from './XlsxPackageContext';

function emptyStructure(descriptor: XlsxSheetDescriptor) {
  return {
    revision: descriptor.revision,
    rowCount: descriptor.rowCount,
    columnCount: descriptor.columnCount,
    defaultRowHeight: 20,
    defaultColumnWidth: 64,
    rows: [],
    columns: [],
    merges: [],
    images: [],
    charts: [],
  };
}

function createEmptyCell(row: number, column: number): SpreadsheetCell {
  return {
    ref: `${columnIndexToLabel(column)}${row}`,
    rowIndex: row,
    columnIndex: column,
    value: '',
  };
}

/** 提供 XLSX 按 Sheet 解析、稀疏范围读取和可重试状态。 */
export class XlsxSpreadsheetSource implements SpreadsheetSource {
  private readonly listeners = new Set<() => void>();
  private readonly descriptors: XlsxSheetDescriptor[];
  private readonly profiles = new Map<string, SpreadsheetPerformanceProfile>();
  private readonly stores = new Map<string, SpreadsheetSheetStore>();
  private readonly requests = new Map<string, Promise<void>>();
  // Sheet 切换只取消调用方等待；底层解析由 Source 生命周期统一管理并继续预热缓存。
  private readonly lifecycleController = new AbortController();
  private revision = 1;
  private snapshot: SpreadsheetSourceSnapshot;
  private disposed = false;
  private disposePromise?: Promise<void>;

  constructor(private readonly context: XlsxPackageContext) {
    this.descriptors = context.descriptors.map((descriptor) => ({
      ...descriptor,
    }));
    this.descriptors.forEach((descriptor) => {
      this.profiles.set(descriptor.id, descriptor.performance);
    });
    this.snapshot = this.createSnapshot();
  }

  private ensureAvailable(sheetId?: string) {
    if (this.disposed) throw new Error('XLSX Source 已释放');
    if (
      sheetId !== undefined &&
      !this.descriptors.some((descriptor) => descriptor.id === sheetId)
    ) {
      throw new RangeError(`工作表不存在：${sheetId}`);
    }
  }

  private getDescriptor(sheetId: string) {
    this.ensureAvailable(sheetId);
    return this.descriptors.find((descriptor) => descriptor.id === sheetId)!;
  }

  private createSnapshot(): SpreadsheetSourceSnapshot {
    return {
      revision: this.revision,
      sheets: this.descriptors.map(
        ({ sheetBytes, relsPath, performance, ...descriptor }) => ({
          ...descriptor,
        }),
      ),
    };
  }

  private publish() {
    this.revision += 1;
    this.snapshot = this.createSnapshot();
    this.listeners.forEach((listener) => listener());
  }

  private updateDescriptor(
    sheetId: string,
    patch: Partial<SpreadsheetSheetDescriptor>,
  ) {
    const index = this.descriptors.findIndex(
      (descriptor) => descriptor.id === sheetId,
    );
    const current = this.descriptors[index];
    this.descriptors[index] = {
      ...current,
      ...patch,
      revision: current.revision + 1,
    };
    this.publish();
  }

  getSnapshot() {
    return this.snapshot;
  }

  subscribe(listener: () => void) {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  getProfile(sheetId: string) {
    this.ensureAvailable(sheetId);
    return this.profiles.get(sheetId)!;
  }

  getSheetLayout(sheetId: string) {
    const descriptor = this.getDescriptor(sheetId);
    return (
      this.stores.get(sheetId)?.getLayout() ?? {
        revision: descriptor.revision,
        rowCount: descriptor.rowCount,
        columnCount: descriptor.columnCount,
        defaultRowHeight: 20,
        defaultColumnWidth: 64,
        rows: [],
        columns: [],
      }
    );
  }

  async ensureSheet(sheetId: string, signal?: AbortSignal) {
    const descriptor = this.getDescriptor(sheetId);
    throwIfSpreadsheetAborted(signal);
    if (descriptor.status === 'ready') return;
    let request = this.requests.get(sheetId);
    if (!request) {
      request = (async () => {
        const taskSignal = this.lifecycleController.signal;
        const startedAt = performance.now();
        let store = this.stores.get(sheetId);
        if (!store) {
          store = createSpreadsheetSheetStore(
            this.context.sessionId,
            sheetId,
            emptyStructure(descriptor),
          );
          this.stores.set(sheetId, store);
        }
        const parsed = await parseXlsxSheetStream(
          this.context,
          descriptor,
          taskSignal,
          {
            collectCells: false,
            onTile: (tile) => store!.putTile(tile),
          },
        );
        const objects = await loadXlsxDrawingObjects(
          this.context,
          descriptor,
          parsed.layout,
          parsed.drawingRelationshipId,
          taskSignal,
        );
        store.putStructure({
          ...parsed.layout,
          merges: parsed.merges,
          images: objects.images,
          charts: objects.charts,
        });
        const elapsed = performance.now() - startedAt;
        const nextProfile = createSpreadsheetPerformanceProfile({
          rowCount: parsed.layout.rowCount,
          columnCount: parsed.layout.columnCount,
          sheetBytes: descriptor.sheetBytes,
          modelBuildMilliseconds: elapsed,
        });
        this.profiles.set(
          sheetId,
          upgradeSpreadsheetPerformanceProfile(
            this.profiles.get(sheetId)!,
            nextProfile,
          ),
        );
        this.updateDescriptor(sheetId, {
          rowCount: parsed.layout.rowCount,
          columnCount: parsed.layout.columnCount,
          status: 'ready',
          errorMessage: undefined,
        });
      })().catch((error) => {
        if (this.lifecycleController.signal.aborted) throw error;
        this.updateDescriptor(sheetId, {
          status: 'error',
          errorMessage:
            error instanceof Error ? error.message : '工作表加载失败',
        });
        throw error;
      });
      this.requests.set(sheetId, request);
      void request.then(
        () => this.requests.delete(sheetId),
        () => this.requests.delete(sheetId),
      );
    }
    await waitForSpreadsheetResult(request, signal);
  }

  async getMaterializedSheet(sheetId: string, signal?: AbortSignal) {
    await this.ensureSheet(sheetId, signal);
    if (this.getProfile(sheetId).gridMode !== 'table') return undefined;
    const descriptor = this.getDescriptor(sheetId);
    const layout = this.getSheetLayout(sheetId);
    const range = await this.getRange(
      sheetId,
      {
        startRow: 1,
        endRow: layout.rowCount,
        startColumn: 1,
        endColumn: layout.columnCount,
      },
      signal,
    );
    const cells = new Map(
      range.cells.map((cell) => [`${cell.rowIndex}:${cell.columnIndex}`, cell]),
    );
    const mergeByPosition = new Map<
      string,
      { rootKey: string; rowSpan: number; colSpan: number }
    >();
    range.merges.forEach((merge) => {
      const rootKey = `${merge.startRow}:${merge.startColumn}`;
      const mergeInfo = {
        rootKey,
        rowSpan: merge.endRow - merge.startRow + 1,
        colSpan: merge.endColumn - merge.startColumn + 1,
      };
      for (let row = merge.startRow; row <= merge.endRow; row += 1) {
        for (
          let column = merge.startColumn;
          column <= merge.endColumn;
          column += 1
        ) {
          mergeByPosition.set(`${row}:${column}`, mergeInfo);
        }
      }
    });
    const rowMetrics = new Map(range.rows.map((row) => [row.index, row]));
    const columnMetrics = new Map(
      range.columns.map((column) => [column.index, column]),
    );
    const sheet: SpreadsheetSheet = {
      id: descriptor.id,
      name: descriptor.name,
      path: descriptor.path,
      kind: descriptor.kind,
      defaultColumnWidth: layout.defaultColumnWidth,
      defaultRowHeight: layout.defaultRowHeight,
      range:
        layout.rowCount === 1 && layout.columnCount === 1
          ? 'A1'
          : `A1:${columnIndexToLabel(layout.columnCount)}${layout.rowCount}`,
      rowCount: layout.rowCount,
      columnCount: layout.columnCount,
      columns: Array.from({ length: layout.columnCount }, (_, offset) => {
        const index = offset + 1;
        const metric = columnMetrics.get(index);
        return {
          index,
          label: columnIndexToLabel(index),
          width: metric?.width ?? layout.defaultColumnWidth,
          hidden: metric?.hidden,
        };
      }),
      rows: Array.from({ length: layout.rowCount }, (_, offset) => {
        const rowIndex = offset + 1;
        const metric = rowMetrics.get(rowIndex);
        return {
          index: rowIndex,
          height: metric?.height ?? layout.defaultRowHeight,
          hidden: metric?.hidden,
          cells: Array.from(
            { length: layout.columnCount },
            (_, columnOffset) => {
              const columnIndex = columnOffset + 1;
              const key = `${rowIndex}:${columnIndex}`;
              const cell =
                cells.get(key) ?? createEmptyCell(rowIndex, columnIndex);
              const merge = mergeByPosition.get(key);
              if (merge && merge.rootKey !== key) {
                return { ...cell, hiddenByMerge: true };
              }
              return merge
                ? {
                    ...cell,
                    rowSpan: merge.rowSpan,
                    colSpan: merge.colSpan,
                  }
                : cell;
            },
          ),
        };
      }),
      merges: [...range.merges],
      images: [...range.images],
      charts: [...range.charts],
    };
    return sheet;
  }

  async getRange(
    sheetId: string,
    range: SpreadsheetRange,
    signal?: AbortSignal,
  ) {
    await this.ensureSheet(sheetId, signal);
    const data = await this.stores.get(sheetId)!.getRange(range, signal);
    const sharedCells = data.cells.filter(
      (cell) => cell.type === 's' && Number.isInteger(Number(cell.value)),
    );
    if (!sharedCells.length) return data;
    const shared = await this.context.sharedStrings.resolveMany(
      sharedCells.map((cell) => Number(cell.value)),
      signal,
    );
    return {
      ...data,
      cells: data.cells.map((cell) =>
        cell.type === 's' && Number.isInteger(Number(cell.value))
          ? {
              ...cell,
              value: shared.get(Number(cell.value)) ?? '',
            }
          : cell,
      ),
    };
  }

  retainRange(sheetId: string, range: SpreadsheetRange) {
    this.ensureAvailable(sheetId);
    return this.stores.get(sheetId)?.retainRange(range) ?? (() => undefined);
  }

  retrySheet(sheetId: string) {
    const descriptor = this.getDescriptor(sheetId);
    if (descriptor.status !== 'error') return;
    const store = this.stores.get(sheetId);
    this.stores.delete(sheetId);
    void store?.dispose();
    this.updateDescriptor(sheetId, {
      status: 'estimated',
      errorMessage: undefined,
    });
  }

  dispose() {
    if (this.disposePromise) return this.disposePromise;
    this.disposed = true;
    this.lifecycleController.abort();
    this.listeners.clear();
    this.disposePromise = Promise.allSettled([...this.requests.values()])
      .then(() =>
        Promise.allSettled([
          ...[...this.stores.values()].map((store) => store.dispose()),
          this.context.sharedStrings.dispose(),
          this.context.reader.close(),
        ]),
      )
      .then(() => undefined);
    return this.disposePromise;
  }
}

/** 从已画像的 XLSX Reader 创建按需工作表 Source。 */
export async function createXlsxSpreadsheetSourceFromArchive(
  archive: ProfiledXlsxArchive,
  sessionId: string,
  signal?: AbortSignal,
) {
  const { context, profile } = await readXlsxStructure(
    archive,
    sessionId,
    signal,
  );
  return {
    source: new XlsxSpreadsheetSource(context),
    profile,
  };
}

/** 打开 XLSX 并创建按需 Source，失败时保证 Reader 被关闭。 */
export async function createXlsxSpreadsheetSource(
  file: File,
  sessionId: string,
  signal?: AbortSignal,
) {
  const archive = await profileXlsxArchive(file, signal);
  try {
    return await createXlsxSpreadsheetSourceFromArchive(
      archive,
      sessionId,
      signal,
    );
  } catch (error) {
    await archive.reader.close();
    throw error;
  }
}
