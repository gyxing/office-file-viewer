import { openOfficeArchive } from '../../shared/ooxml/archive';
import type { OfficeArchiveReader } from '../../shared/ooxml/OfficeArchiveReader';
import { ResourceRegistry } from '../parsing/assembly/ResourceRegistry';
import { createSpreadsheetAxisIndex } from '../spreadsheet/SpreadsheetAxisIndex';
import {
  createSpreadsheetPerformanceProfile,
  upgradeSpreadsheetPerformanceProfile,
  type SpreadsheetPerformanceProfile,
} from '../spreadsheet/spreadsheetPerformance';
import {
  createSpreadsheetSheetStore,
  SPREADSHEET_TILE_COLUMNS,
  SPREADSHEET_TILE_ROWS,
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
  SpreadsheetChart,
  SpreadsheetImage,
  SpreadsheetRange,
  SpreadsheetSheet,
} from '../spreadsheet/types';
import {
  loadWpsCellImages,
  readWpsCellImagePlacement,
} from '../spreadsheet/wpsCellImages';
import { adaptBiff8WorksheetSparse } from './adapter';
import { BIFF8_RECORD } from './biff8/constants';
import { parseBiff8WorksheetChunks } from './biff8/parseWorksheetChunks';
import { parseBiff8Charts, readBiff8ChartSubstream } from './chart/parseCharts';
import { createPortableImageResource } from './drawing/createPortableImageResource';
import {
  parseBiff8Drawings,
  parseBiff8DrawingShapes,
} from './drawing/parseDrawings';
import {
  profileXlsArchive,
  readXlsStructure,
  type ProfiledXlsArchive,
  type XlsSheetDescriptor,
  type XlsStructure,
} from './readXlsStructure';
import type { Biff8Workbook, Biff8Worksheet } from './types';

function columnLabel(index: number) {
  let value = index;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function concatChunks(chunks: readonly Uint8Array[]) {
  const result = new Uint8Array(
    chunks.reduce((total, chunk) => total + chunk.length, 0),
  );
  let offset = 0;
  chunks.forEach((chunk) => {
    result.set(chunk, offset);
    offset += chunk.length;
  });
  return result;
}

function emptyStructure(descriptor: XlsSheetDescriptor) {
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

/** 根据稀疏行列尺寸把 BIFF8 ClientAnchor 转为完整工作表坐标。 */
function anchorGeometry(
  anchor: {
    from: {
      row: number;
      column: number;
      rowFraction: number;
      columnFraction: number;
    };
    to: {
      row: number;
      column: number;
      rowFraction: number;
      columnFraction: number;
    };
  },
  sparse: ReturnType<typeof adaptBiff8WorksheetSparse>,
) {
  const rowAxis = createSpreadsheetAxisIndex(
    sparse.rowCount,
    sparse.defaultRowHeight,
    sparse.rows,
  );
  const columnAxis = createSpreadsheetAxisIndex(
    sparse.columnCount,
    sparse.defaultColumnWidth,
    sparse.columns,
  );
  const fromRow = anchor.from.row + 1;
  const fromColumn = anchor.from.column + 1;
  const toRow = anchor.to.row + 1;
  const toColumn = anchor.to.column + 1;
  const x =
    columnAxis.offsetAt(fromColumn) +
    columnAxis.sizeAt(fromColumn) * anchor.from.columnFraction;
  const y =
    rowAxis.offsetAt(fromRow) +
    rowAxis.sizeAt(fromRow) * anchor.from.rowFraction;
  const right =
    columnAxis.offsetAt(toColumn) +
    columnAxis.sizeAt(toColumn) * anchor.to.columnFraction;
  const bottom =
    rowAxis.offsetAt(toRow) + rowAxis.sizeAt(toRow) * anchor.to.rowFraction;
  return {
    from: {
      row: fromRow,
      column: fromColumn,
      rowOffset: rowAxis.sizeAt(fromRow) * anchor.from.rowFraction,
      columnOffset: columnAxis.sizeAt(fromColumn) * anchor.from.columnFraction,
    },
    to: {
      row: toRow,
      column: toColumn,
      rowOffset: rowAxis.sizeAt(toRow) * anchor.to.rowFraction,
      columnOffset: columnAxis.sizeAt(toColumn) * anchor.to.columnFraction,
    },
    x,
    y,
    width: Math.max(1, right - x),
    height: Math.max(1, bottom - y),
  };
}

/** 解析当前 XLS Sheet 的图片和图表，并保持全局锚点坐标。 */
async function loadXlsObjects(
  structure: XlsStructure,
  workbook: Biff8Workbook,
  worksheet: Biff8Worksheet,
  sparse: ReturnType<typeof adaptBiff8WorksheetSparse>,
  resources: ResourceRegistry,
) {
  const groupBytes = concatChunks(
    structure.globals.drawingGroupRecords.flatMap((record) => record.chunks),
  );
  const drawingBytes = concatChunks(
    worksheet.drawingRecords
      .filter((record) => record.recordId === BIFF8_RECORD.MSODRAWING)
      .flatMap((record) => record.chunks),
  );
  const warnings = workbook.warnings;
  const images: SpreadsheetImage[] = [];
  if (groupBytes.length && drawingBytes.length) {
    const drawings = parseBiff8Drawings(groupBytes, drawingBytes, warnings);
    for (let index = 0; index < drawings.length; index += 1) {
      const drawing = drawings[index];
      const portable = await createPortableImageResource(
        drawing,
        `xls-source:${worksheet.descriptor.id}:${drawing.id}:${index}`,
      );
      const src = await resources.register(portable.resource);
      images.push({
        id: drawing.id,
        name: drawing.name,
        alt: drawing.alt,
        src,
        ...anchorGeometry(drawing.anchor, sparse),
      });
    }
  }

  const charts: SpreadsheetChart[] = [];
  if (worksheet.chartSubstreams.length) {
    const shapes = parseBiff8DrawingShapes(drawingBytes, warnings);
    parseBiff8Charts(
      workbook,
      worksheet.descriptor,
      worksheet.chartSubstreams,
      shapes,
      images,
      worksheet,
    ).forEach((item) => {
      charts.push({
        id: item.id,
        title: item.title,
        chart: item.chart,
        ...anchorGeometry(item.anchor, sparse),
      });
    });
  }
  return { images, charts };
}

function groupTiles(
  sheetId: string,
  revision: number,
  cells: readonly SpreadsheetCell[],
) {
  const tiles = new Map<string, SpreadsheetCell[]>();
  cells.forEach((cell) => {
    const rowTile = Math.floor((cell.rowIndex - 1) / SPREADSHEET_TILE_ROWS);
    const columnTile = Math.floor(
      (cell.columnIndex - 1) / SPREADSHEET_TILE_COLUMNS,
    );
    const key = `${rowTile}:${columnTile}`;
    const bucket = tiles.get(key) ?? [];
    bucket.push(cell);
    tiles.set(key, bucket);
  });
  return [...tiles].map(([key, tileCells]) => {
    const [rowTile, columnTile] = key.split(':').map(Number);
    return {
      sheetId,
      rowTile,
      columnTile,
      revision,
      cells: tileCells,
    };
  });
}

/** 只读取并适配当前独立 Chart Sheet，不触碰其他工作表正文。 */
async function loadXlsChartSheet(
  structure: XlsStructure,
  descriptor: XlsSheetDescriptor,
  signal: AbortSignal,
) {
  const bytes = await structure.workbookStream.read(
    descriptor.streamOffset,
    Math.max(0, descriptor.endOffset - descriptor.streamOffset),
    signal,
  );
  const localDescriptor = { ...descriptor, streamOffset: 0 };
  const substream = readBiff8ChartSubstream(
    bytes,
    localDescriptor,
    bytes.length,
  );
  const workbook: Biff8Workbook = {
    globals: structure.globals,
    worksheets: [],
    chartSheets: [{ descriptor, substream }],
    warnings: [],
  };
  const parsed = parseBiff8Charts(workbook, descriptor, [substream], [], [])[0];
  if (!parsed) return [];
  return [
    {
      id: parsed.id,
      title: parsed.title,
      chart: parsed.chart,
      from: { row: 1, column: 1, rowOffset: 0, columnOffset: 0 },
      to: { row: 1, column: 1, rowOffset: 0, columnOffset: 0 },
      x: 0,
      y: 0,
      width: 960,
      height: 600,
    },
  ] satisfies SpreadsheetChart[];
}

/** 提供 XLS 的 CFB 随机访问和按 Sheet 解析。 */
export class XlsSpreadsheetSource implements SpreadsheetSource {
  private readonly listeners = new Set<() => void>();
  private readonly descriptors: XlsSheetDescriptor[];
  private readonly profiles = new Map<string, SpreadsheetPerformanceProfile>();
  private readonly stores = new Map<string, SpreadsheetSheetStore>();
  private readonly requests = new Map<string, Promise<void>>();
  private readonly resources = new ResourceRegistry();
  private cellImageArchive?: OfficeArchiveReader;
  private cellImageArchivePromise?: Promise<OfficeArchiveReader | undefined>;
  // Sheet 切换只取消调用方等待；底层解析由 Source 生命周期统一管理并继续预热缓存。
  private readonly lifecycleController = new AbortController();
  private revision = 1;
  private snapshot: SpreadsheetSourceSnapshot;
  private disposed = false;
  private disposePromise?: Promise<void>;

  constructor(private readonly structure: XlsStructure) {
    this.descriptors = structure.descriptors.map((descriptor) => ({
      ...descriptor,
    }));
    this.descriptors.forEach((descriptor) =>
      this.profiles.set(descriptor.id, descriptor.performance),
    );
    this.snapshot = this.createSnapshot();
  }

  private ensureAvailable(sheetId?: string) {
    if (this.disposed) throw new Error('XLS Source 已释放');
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

  private getCellImageArchive(signal: AbortSignal) {
    if (!this.cellImageArchivePromise) {
      this.cellImageArchivePromise = (async () => {
        const stream = this.structure.reader.openStream('ETCellImageData');
        if (!stream) return undefined;
        const reader = await openOfficeArchive(
          await stream.materialize(signal),
          { signal },
        );
        this.cellImageArchive = reader;
        return reader;
      })().catch(() => undefined);
    }
    return this.cellImageArchivePromise;
  }

  private createSnapshot(): SpreadsheetSourceSnapshot {
    return {
      revision: this.revision,
      sheets: this.descriptors.map(
        ({
          endOffset,
          streamOffset,
          performance,
          type,
          visibility,
          ...item
        }) => ({
          ...item,
          path: `/Workbook/${item.name}`,
          kind: type === 'chart' ? 'chart' : 'worksheet',
        }),
      ),
    };
  }

  private updateDescriptor(
    sheetId: string,
    patch: Partial<SpreadsheetSheetDescriptor>,
  ) {
    const index = this.descriptors.findIndex((item) => item.id === sheetId);
    const current = this.descriptors[index];
    this.descriptors[index] = {
      ...current,
      ...patch,
      revision: current.revision + 1,
    };
    this.revision += 1;
    this.snapshot = this.createSnapshot();
    this.listeners.forEach((listener) => listener());
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
            this.structure.sessionId,
            sheetId,
            emptyStructure(descriptor),
          );
          this.stores.set(sheetId, store);
        }
        if (descriptor.type === 'chart') {
          const charts = await loadXlsChartSheet(
            this.structure,
            descriptor,
            taskSignal,
          );
          store.putStructure({
            ...emptyStructure(descriptor),
            revision: descriptor.revision + 1,
            charts,
          });
          this.updateDescriptor(sheetId, { status: 'ready' });
          return;
        }
        if (descriptor.type !== 'worksheet') {
          store.putStructure({
            ...emptyStructure(descriptor),
            revision: descriptor.revision + 1,
          });
          this.updateDescriptor(sheetId, { status: 'ready' });
          return;
        }
        const worksheet = await parseBiff8WorksheetChunks(
          this.structure.workbookStream,
          descriptor,
          descriptor.endOffset,
          this.structure.globals,
          taskSignal,
        );
        const workbook: Biff8Workbook = {
          globals: this.structure.globals,
          worksheets: [worksheet],
          chartSheets: [],
          warnings: [...worksheet.warnings],
        };
        const sparse = adaptBiff8WorksheetSparse(workbook, worksheet);
        const cellImagePlacements = sparse.cells.flatMap((cell) => {
          const placement = readWpsCellImagePlacement(cell);
          if (!placement) return [];
          // DISPIMG 的缓存值只是兼容公式文本；成功与否都不应作为正文显示。
          cell.value = '';
          return [placement];
        });
        const objects = await loadXlsObjects(
          this.structure,
          workbook,
          worksheet,
          sparse,
          this.resources,
        );
        const cellImageArchive = cellImagePlacements.length
          ? await this.getCellImageArchive(taskSignal)
          : undefined;
        const cellImages = cellImageArchive
          ? await loadWpsCellImages(
              cellImageArchive,
              this.structure.sessionId,
              cellImagePlacements,
              sparse,
              sparse.merges,
              taskSignal,
            )
          : [];
        const revision = descriptor.revision + 1;
        store.putStructure({
          revision,
          rowCount: sparse.rowCount,
          columnCount: sparse.columnCount,
          defaultRowHeight: sparse.defaultRowHeight,
          defaultColumnWidth: sparse.defaultColumnWidth,
          rows: sparse.rows,
          columns: sparse.columns,
          merges: sparse.merges,
          images: [...objects.images, ...cellImages],
          charts: objects.charts,
        });
        await Promise.all(
          groupTiles(sheetId, revision, sparse.cells).map((tile) =>
            store!.putTile(tile),
          ),
        );
        const nextProfile = createSpreadsheetPerformanceProfile({
          rowCount: sparse.rowCount,
          columnCount: sparse.columnCount,
          cfbFileBytes: this.structure.fileSize,
          cfbMainStreamBytes: this.structure.mainStreamSize,
          sheetBytes: descriptor.endOffset - descriptor.streamOffset,
          modelBuildMilliseconds: performance.now() - startedAt,
        });
        this.profiles.set(
          sheetId,
          upgradeSpreadsheetPerformanceProfile(
            this.profiles.get(sheetId)!,
            nextProfile,
          ),
        );
        this.updateDescriptor(sheetId, {
          rowCount: sparse.rowCount,
          columnCount: sparse.columnCount,
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
    const data = await this.getRange(
      sheetId,
      {
        startRow: 1,
        endRow: layout.rowCount,
        startColumn: 1,
        endColumn: layout.columnCount,
      },
      signal,
    );
    const cellMap = new Map(
      data.cells.map((cell) => [`${cell.rowIndex}:${cell.columnIndex}`, cell]),
    );
    const mergeRoots = new Map<string, string>();
    data.merges.forEach((merge) => {
      const root = `${merge.startRow}:${merge.startColumn}`;
      for (let row = merge.startRow; row <= merge.endRow; row += 1) {
        for (
          let column = merge.startColumn;
          column <= merge.endColumn;
          column += 1
        ) {
          mergeRoots.set(`${row}:${column}`, root);
        }
      }
    });
    const rowMap = new Map(data.rows.map((row) => [row.index, row]));
    const columnMap = new Map(
      data.columns.map((column) => [column.index, column]),
    );
    const sheet: SpreadsheetSheet = {
      id: descriptor.id,
      name: descriptor.name,
      path: `/Workbook/${descriptor.name}`,
      kind: descriptor.type === 'chart' ? 'chart' : 'worksheet',
      defaultRowHeight: layout.defaultRowHeight,
      defaultColumnWidth: layout.defaultColumnWidth,
      range:
        layout.rowCount === 1 && layout.columnCount === 1
          ? 'A1'
          : `A1:${columnLabel(layout.columnCount)}${layout.rowCount}`,
      rowCount: layout.rowCount,
      columnCount: layout.columnCount,
      columns: Array.from({ length: layout.columnCount }, (_, offset) => {
        const index = offset + 1;
        const metric = columnMap.get(index);
        return {
          index,
          label: columnLabel(index),
          width: metric?.width ?? layout.defaultColumnWidth,
          hidden: metric?.hidden,
        };
      }),
      rows: Array.from({ length: layout.rowCount }, (_, offset) => {
        const rowIndex = offset + 1;
        const metric = rowMap.get(rowIndex);
        return {
          index: rowIndex,
          height: metric?.height ?? layout.defaultRowHeight,
          hidden: metric?.hidden,
          cells: Array.from(
            { length: layout.columnCount },
            (_, columnOffset) => {
              const columnIndex = columnOffset + 1;
              const key = `${rowIndex}:${columnIndex}`;
              const cell = cellMap.get(key) ?? {
                ref: `${columnLabel(columnIndex)}${rowIndex}`,
                rowIndex,
                columnIndex,
                value: '',
              };
              const root = mergeRoots.get(key);
              return root && root !== key
                ? { ...cell, hiddenByMerge: true }
                : cell;
            },
          ),
        };
      }),
      merges: [...data.merges],
      images: [...data.images],
      charts: [...data.charts],
    };
    return sheet;
  }

  async getRange(
    sheetId: string,
    range: SpreadsheetRange,
    signal?: AbortSignal,
  ) {
    await this.ensureSheet(sheetId, signal);
    return this.stores.get(sheetId)!.getRange(range, signal);
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
          this.structure.sharedStrings.dispose(),
          this.cellImageArchive?.close(),
          this.structure.reader.close(),
        ]),
      )
      .then(() => {
        this.resources.dispose();
      });
    return this.disposePromise;
  }
}

/** 从已打开的 CFB Reader 创建 XLS 按需 Source。 */
export async function createXlsSpreadsheetSourceFromArchive(
  archive: ProfiledXlsArchive,
  sessionId: string,
  signal?: AbortSignal,
) {
  const result = await readXlsStructure(archive, sessionId, signal);
  return {
    source: new XlsSpreadsheetSource(result.structure),
    requiresSource: result.requiresSource,
  };
}

/** 打开 XLS 并创建按需 Source，失败时保证 CFB Reader 被关闭。 */
export async function createXlsSpreadsheetSource(
  file: File,
  sessionId: string,
  signal?: AbortSignal,
) {
  const archive = await profileXlsArchive(file, signal);
  try {
    return await createXlsSpreadsheetSourceFromArchive(
      archive,
      sessionId,
      signal,
    );
  } catch (error) {
    await archive.reader.close();
    throw error;
  }
}
