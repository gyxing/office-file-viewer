import { readOfficeXmlEvents } from '../../shared/ooxml/OfficeXmlEventReader';
import {
  SPREADSHEET_TILE_COLUMNS,
  SPREADSHEET_TILE_ROWS,
  type SpreadsheetTile,
} from '../spreadsheet/SpreadsheetSheetStore';
import {
  throwIfSpreadsheetAborted,
  type SpreadsheetSheetLayout,
} from '../spreadsheet/SpreadsheetSource';
import type {
  SpreadsheetCell,
  SpreadsheetColumnMetric,
  SpreadsheetMerge,
  SpreadsheetRowMetric,
} from '../spreadsheet/types';
import {
  readWpsCellImageId,
  type WpsCellImagePlacement,
} from '../spreadsheet/wpsCellImages';
import type {
  XlsxPackageContext,
  XlsxSheetDescriptor,
} from './XlsxPackageContext';
import {
  columnIndexToLabel,
  excelWidthToPx,
  parseCellRef,
  parseRange,
  pointToPx,
  resolveStyle,
  resolveXlsxMaxDigitWidth,
} from './xlsxCellFormatting';

/** 单个 Sheet 流式解析后的稀疏结构。 */
export type ParsedXlsxSheetStream = {
  /** 当前内容使用的布局信息。 */
  layout: SpreadsheetSheetLayout;
  /** 按显示顺序排列的单元格。 */
  cells: readonly SpreadsheetCell[];
  /** 当前工作表声明的合并区域。 */
  merges: readonly SpreadsheetMerge[];
  /** 流式解析过程中生成的工作表分片。 */
  tiles: readonly SpreadsheetTile[];
  /** 由 WPS 单元格图片公式引用的图片位置。 */
  cellImages: readonly WpsCellImagePlacement[];
  /** 当前工作表绘图部件的关系标识。 */
  drawingRelationshipId?: string;
};

/** 控制 Sheet 解析期间的 tile 输出，避免大 Sheet 聚合全部单元格。 */
export type ParseXlsxSheetStreamOptions = {
  /** 是否在输出中保留全部单元格；分片模式可关闭以降低内存占用。 */
  collectCells?: boolean;
  /** 分片发生时触发的回调。 */
  onTile?: (tile: SpreadsheetTile) => Promise<void>;
};

/** XML 事件读取期间尚未完成组装的 XLSX 单元格。 */
type PendingCell = {
  /** 待解析单元格的 A1 引用。 */
  ref: string;
  /** 行的零基索引。 */
  rowIndex: number;
  /** 列的零基索引。 */
  columnIndex: number;
  /** 样式标识。 */
  styleId?: number;
  /** OOXML 单元格值的编码类型。 */
  type?: string;
  /** 单元格值节点累积的原始文本。 */
  value: string;
  /** 单元格公式节点累积的公式文本。 */
  formula: string;
  /** 单元格内联字符串累积的文本。 */
  inlineText: string;
};

/** 等待转换为标准列宽模型的 XLSX 列区间。 */
type PendingColumnMetric = {
  /** 当前范围的起始位置。 */
  start: number;
  /** 当前范围的结束位置。 */
  end: number;
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 是否隐藏当前项目。 */
  hidden: boolean;
};

function readBoolean(value: string | undefined) {
  return value === '1' || value === 'true';
}

function tileCoordinates(cell: SpreadsheetCell) {
  return {
    rowTile: Math.floor((cell.rowIndex - 1) / SPREADSHEET_TILE_ROWS),
    columnTile: Math.floor((cell.columnIndex - 1) / SPREADSHEET_TILE_COLUMNS),
  };
}

/** 流式读取一个 worksheet，只保存实际单元格和稀疏尺寸覆盖。 */
export async function parseXlsxSheetStream(
  context: XlsxPackageContext,
  descriptor: XlsxSheetDescriptor,
  signal?: AbortSignal,
  options: ParseXlsxSheetStreamOptions = {},
): Promise<ParsedXlsxSheetStream> {
  throwIfSpreadsheetAborted(signal);
  const stream = await context.reader.openStream(descriptor.path, signal);
  const collectCells = options.collectCells ?? !options.onTile;
  const cells: SpreadsheetCell[] = [];
  const pendingTiles = new Map<string, SpreadsheetCell[]>();
  const rows = new Map<number, SpreadsheetRowMetric>();
  const columns = new Map<number, SpreadsheetColumnMetric>();
  const pendingColumns: PendingColumnMetric[] = [];
  const merges: SpreadsheetMerge[] = [];
  const cellImages: WpsCellImagePlacement[] = [];
  let rowCount = Math.max(1, descriptor.rowCount);
  let columnCount = Math.max(1, descriptor.columnCount);
  let defaultRowHeight = 20;
  let defaultColumnWidth = 64;
  const maxDigitWidth = resolveXlsxMaxDigitWidth(context.styles.fonts[0]);
  let pendingCell: PendingCell | undefined;
  let capture: 'value' | 'formula' | 'inline' | undefined;
  let drawingRelationshipId: string | undefined;
  let activeRowTile = 0;

  const flushTiles = async () => {
    if (!pendingTiles.size) return;
    const revision = descriptor.revision + 1;
    const tiles = [...pendingTiles].map(([key, tileCells]) => {
      const [rowTile, columnTile] = key.split(':').map(Number);
      return {
        sheetId: descriptor.id,
        rowTile,
        columnTile,
        revision,
        cells: tileCells,
      } satisfies SpreadsheetTile;
    });
    pendingTiles.clear();
    if (options.onTile) {
      await Promise.all(tiles.map((tile) => options.onTile!(tile)));
    }
  };

  for await (const event of readOfficeXmlEvents(stream, signal)) {
    throwIfSpreadsheetAborted(signal);
    if (event.type === 'open') {
      if (event.localName === 'dimension') {
        const range = parseRange(event.attributes.get('ref'));
        if (range) {
          rowCount = Math.max(rowCount, range.endRow);
          columnCount = Math.max(columnCount, range.endColumn);
        }
      } else if (event.localName === 'sheetFormatPr') {
        const sourceDefaultColumnWidth =
          event.attributes.get('defaultColWidth');
        defaultColumnWidth = sourceDefaultColumnWidth
          ? excelWidthToPx(
              Number(sourceDefaultColumnWidth),
              defaultColumnWidth,
              maxDigitWidth,
            )
          : defaultColumnWidth;
        defaultRowHeight = pointToPx(
          Number(event.attributes.get('defaultRowHeight') ?? 15),
        );
      } else if (event.localName === 'col') {
        const start = Math.max(1, Number(event.attributes.get('min') ?? 1));
        const end = Math.max(
          start,
          Number(event.attributes.get('max') ?? start),
        );
        const width = excelWidthToPx(
          Number(event.attributes.get('width')),
          defaultColumnWidth,
          maxDigitWidth,
        );
        const hidden = readBoolean(event.attributes.get('hidden'));
        pendingColumns.push({ start, end, width, hidden });
      } else if (event.localName === 'row') {
        const index = Math.max(1, Number(event.attributes.get('r') ?? 1));
        const hasHeight = event.attributes.has('ht');
        const customHeight = readBoolean(event.attributes.get('customHeight'));
        const hidden = readBoolean(event.attributes.get('hidden'));
        if (hasHeight || customHeight || hidden) {
          rows.set(index, {
            index,
            height: hasHeight
              ? pointToPx(Number(event.attributes.get('ht')), defaultRowHeight)
              : defaultRowHeight,
            customHeight,
            hidden,
          });
        }
      } else if (event.localName === 'c') {
        const ref = event.attributes.get('r') ?? 'A1';
        const address = parseCellRef(ref);
        const rawStyleId = event.attributes.get('s');
        pendingCell = {
          ref,
          rowIndex: address.row,
          columnIndex: address.column,
          styleId: rawStyleId === undefined ? undefined : Number(rawStyleId),
          type: event.attributes.get('t'),
          value: '',
          formula: '',
          inlineText: '',
        };
        rowCount = Math.max(rowCount, address.row);
        columnCount = Math.max(columnCount, address.column);
      } else if (pendingCell && event.localName === 'v') {
        capture = 'value';
      } else if (pendingCell && event.localName === 'f') {
        capture = 'formula';
      } else if (
        pendingCell &&
        pendingCell.type === 'inlineStr' &&
        event.localName === 't'
      ) {
        capture = 'inline';
      } else if (event.localName === 'mergeCell') {
        const ref = event.attributes.get('ref') ?? '';
        const range = parseRange(ref);
        if (range) {
          merges.push({ ref, ...range });
          rowCount = Math.max(rowCount, range.endRow);
          columnCount = Math.max(columnCount, range.endColumn);
        }
      } else if (event.localName === 'drawing') {
        drawingRelationshipId =
          event.attributes.get('r:id') ?? event.attributes.get('id');
      }
      continue;
    }

    if (event.type === 'text' && pendingCell && capture) {
      if (capture === 'value') pendingCell.value += event.text;
      else if (capture === 'formula') pendingCell.formula += event.text;
      else pendingCell.inlineText += event.text;
      continue;
    }

    if (event.type !== 'close') continue;
    if (
      event.localName === 'v' ||
      event.localName === 'f' ||
      event.localName === 't'
    ) {
      capture = undefined;
    }
    if (event.localName !== 'c' || !pendingCell) continue;

    const rawValue = pendingCell.value;
    const cellImageId = readWpsCellImageId(pendingCell.formula, rawValue);
    const cell: SpreadsheetCell = {
      ref: pendingCell.ref,
      rowIndex: pendingCell.rowIndex,
      columnIndex: pendingCell.columnIndex,
      rawValue,
      value: cellImageId
        ? ''
        : pendingCell.type === 'inlineStr'
        ? pendingCell.inlineText
        : pendingCell.type === 'b'
        ? rawValue === '1'
          ? 'TRUE'
          : 'FALSE'
        : rawValue,
      type: pendingCell.type,
      styleId: pendingCell.styleId,
      style: resolveStyle(pendingCell.styleId, context.styles),
      formula: pendingCell.formula || undefined,
    };
    if (cellImageId) {
      cellImages.push({
        imageId: cellImageId,
        row: pendingCell.rowIndex,
        column: pendingCell.columnIndex,
      });
    }
    const { rowTile, columnTile } = tileCoordinates(cell);
    if (pendingTiles.size && rowTile !== activeRowTile) {
      await flushTiles();
    }
    activeRowTile = rowTile;
    const tileKey = `${rowTile}:${columnTile}`;
    const tileCells = pendingTiles.get(tileKey) ?? [];
    tileCells.push(cell);
    pendingTiles.set(tileKey, tileCells);
    if (collectCells) cells.push(cell);
    pendingCell = undefined;
    capture = undefined;
  }

  await flushTiles();

  // 列格式可能延伸到 XFD；内容边界确定后再截取，避免把样式范围当成已用列。
  pendingColumns.forEach(({ start, end, width, hidden }) => {
    for (let index = start; index <= Math.min(end, columnCount); index += 1) {
      columns.set(index, { index, width, hidden });
    }
  });

  if (collectCells) {
    const sharedCells = cells
      .filter(
        (cell) => cell.type === 's' && Number.isInteger(Number(cell.value)),
      )
      .map((cell) => ({ cell, index: Number(cell.value) }));
    const shared = await context.sharedStrings.resolveMany(
      sharedCells.map((item) => item.index),
      signal,
    );
    sharedCells.forEach(({ cell, index }) => {
      cell.value = shared.get(index) ?? '';
    });

    const cellByRef = new Map(cells.map((cell) => [cell.ref, cell]));
    merges.forEach((merge) => {
      const ref = `${columnIndexToLabel(merge.startColumn)}${merge.startRow}`;
      let root = cellByRef.get(ref);
      if (!root) {
        root = {
          ref,
          rowIndex: merge.startRow,
          columnIndex: merge.startColumn,
          value: '',
        };
        cellByRef.set(ref, root);
        cells.push(root);
      }
      root.rowSpan = merge.endRow - merge.startRow + 1;
      root.colSpan = merge.endColumn - merge.startColumn + 1;
    });
  }

  const collectedTiles = new Map<string, SpreadsheetCell[]>();
  if (!options.onTile) {
    cells.forEach((cell) => {
      const { rowTile, columnTile } = tileCoordinates(cell);
      const key = `${rowTile}:${columnTile}`;
      const tileCells = collectedTiles.get(key) ?? [];
      tileCells.push(cell);
      collectedTiles.set(key, tileCells);
    });
  }
  const revision = descriptor.revision + 1;
  return {
    layout: {
      revision,
      rowCount,
      columnCount,
      defaultRowHeight,
      defaultColumnWidth,
      rows: [...rows.values()].filter((row) => row.index <= rowCount),
      columns: [...columns.values()],
    },
    cells,
    merges,
    cellImages,
    tiles: [...collectedTiles].map(([key, tileCells]) => {
      const [rowTile, columnTile] = key.split(':').map(Number);
      return {
        sheetId: descriptor.id,
        rowTile,
        columnTile,
        revision,
        cells: tileCells,
      };
    }),
    drawingRelationshipId,
  };
}
