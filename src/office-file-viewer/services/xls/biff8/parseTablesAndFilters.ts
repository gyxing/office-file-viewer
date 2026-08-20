import type {
  SpreadsheetAutoFilter,
  SpreadsheetTable,
} from '../../spreadsheet/semantics/types';
import type {
  SpreadsheetRange,
  SpreadsheetWarning,
} from '../../spreadsheet/types';
import type { Biff8SheetDescriptor, Biff8WorkbookGlobals } from '../types';

function readUint16(bytes: Uint8Array, offset: number) {
  if (offset + 2 > bytes.length) return undefined;
  return new DataView(
    bytes.buffer,
    bytes.byteOffset,
    bytes.byteLength,
  ).getUint16(offset, true);
}

function columnLabel(column: number) {
  let value = Math.max(1, Math.trunc(column));
  let label = '';
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return label;
}

function rangeRef(range: SpreadsheetRange) {
  return `${columnLabel(range.startColumn)}${range.startRow}:${columnLabel(
    range.endColumn,
  )}${range.endRow}`;
}

/** 从 _FilterDatabase 的单个 PtgArea3d 恢复 AutoFilter 范围。 */
export function readBiff8AutoFilterRange(
  globals: Biff8WorkbookGlobals,
  descriptor: Biff8SheetDescriptor,
) {
  const sheetIndex = globals.sheets.findIndex(
    (sheet) => sheet.id === descriptor.id,
  );
  const definedName = globals.definedNames.find(
    (item) =>
      item.builtInId === 0x0d &&
      (item.sheetIndex === undefined || item.sheetIndex === sheetIndex + 1),
  );
  const tokens = definedName?.tokens;
  if (!tokens || tokens.length < 11 || (tokens[0] & 0x1f) !== 0x1b) {
    return undefined;
  }
  const view = new DataView(
    tokens.buffer,
    tokens.byteOffset,
    tokens.byteLength,
  );
  const externalIndex = view.getUint16(1, true);
  const external = globals.externalSheets[externalIndex];
  if (
    external?.firstSheetIndex !== undefined &&
    external.firstSheetIndex !== sheetIndex
  ) {
    return undefined;
  }
  return {
    startRow: view.getUint16(3, true) + 1,
    endRow: view.getUint16(5, true) + 1,
    startColumn: (view.getUint16(7, true) & 0x00ff) + 1,
    endColumn: (view.getUint16(9, true) & 0x00ff) + 1,
  } satisfies SpreadsheetRange;
}

/** 判断单个 AutoFilter 记录是否保存了实际筛选条件。 */
export function readBiff8FilteredColumn(bytes: Uint8Array) {
  if (bytes.length < 24) return undefined;
  const column = readUint16(bytes, 0);
  const flags = readUint16(bytes, 2) ?? 0;
  const hasCondition =
    bytes[4] !== 0 || bytes[14] !== 0 || Boolean(flags & 0x0010);
  return hasCondition ? column : undefined;
}

/** 组装 BIFF8 AutoFilter 的范围和源筛选字段。 */
export function buildBiff8AutoFilter(options: {
  /** 当前工作表。 */
  descriptor: Biff8SheetDescriptor;
  /** 工作簿级定义名称和外部表映射。 */
  globals: Biff8WorkbookGlobals;
  /** AutoFilterInfo 声明的字段数。 */
  entryCount?: number;
  /** AutoFilter 记录中实际带条件的字段索引。 */
  filteredColumns: readonly number[];
  /** DIMENSIONS 提供的回退范围。 */
  fallbackRange?: SpreadsheetRange;
  /** 接收局部降级提示。 */
  warnings: SpreadsheetWarning[];
}): SpreadsheetAutoFilter | undefined {
  if (!options.entryCount) return undefined;
  let range = readBiff8AutoFilterRange(options.globals, options.descriptor);
  if (!range && options.fallbackRange) {
    range = {
      ...options.fallbackRange,
      endColumn: Math.min(
        options.fallbackRange.endColumn,
        options.fallbackRange.startColumn + options.entryCount - 1,
      ),
    };
    options.warnings.push({
      code: 'XLS_AUTOFILTER_RANGE_DEGRADED',
      message: '未能读取 _FilterDatabase，已按工作表已用范围恢复筛选表头。',
      sheetName: options.descriptor.name,
    });
  }
  if (!range) return undefined;
  return {
    ref: rangeRef(range),
    range,
    filteredColumns: [...new Set(options.filteredColumns)].sort(
      (left, right) => left - right,
    ),
  };
}

/** 从 Feature11/12 的 refs2 恢复旧版 Excel Table 基础范围。 */
export function parseBiff8FeatureTable(
  bytes: Uint8Array,
  descriptor: Biff8SheetDescriptor,
  tableIndex: number,
  warnings: SpreadsheetWarning[],
): SpreadsheetTable | undefined {
  if (bytes.length < 35 || readUint16(bytes, 12) !== 0x0005) return undefined;
  const rangeCount = readUint16(bytes, 19) ?? 0;
  if (!rangeCount || 27 + rangeCount * 8 > bytes.length) return undefined;
  const range = {
    startRow: (readUint16(bytes, 27) ?? 0) + 1,
    endRow: (readUint16(bytes, 29) ?? 0) + 1,
    startColumn: (readUint16(bytes, 31) ?? 0) + 1,
    endColumn: (readUint16(bytes, 33) ?? 0) + 1,
  };
  if (range.startRow > range.endRow || range.startColumn > range.endColumn) {
    return undefined;
  }
  warnings.push({
    code: 'XLS_TABLE_STYLE_PARTIAL',
    message: '已恢复 BIFF8 Table 范围和表头，复杂 List12 样式按静态语义降级。',
    sheetName: descriptor.name,
  });
  return {
    id: `${descriptor.id}:table:${tableIndex}`,
    name: `Table${tableIndex}`,
    ref: rangeRef(range),
    range,
    headerRow: true,
    totalsRow: false,
  };
}
