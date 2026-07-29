import { OFFICE_LARGE_FILE_THRESHOLDS } from '../performance/officePerformanceThresholds';

/** 大工作表启用行窗口的可见行边界。 */
export const SPREADSHEET_ROW_WINDOW_THRESHOLD = 500;
/** 大工作表启用二维窗口的可见列边界。 */
export const SPREADSHEET_COLUMN_WINDOW_THRESHOLD = 100;
/** 工作表投影槽位启用窗口化的边界。 */
export const SPREADSHEET_PROJECTED_SLOT_THRESHOLD = 20_000;

/** 描述单个工作表的解析和渲染性能模式。 */
export type SpreadsheetPerformanceProfile = {
  /** 当前 Sheet 是否只在激活时解析。 */
  sheetMode: 'materialized' | 'lazy';
  /** 当前 Sheet 使用完整表格、行窗口或二维窗口。 */
  gridMode: 'table' | 'row-window' | 'two-axis-window';
  /** 行列相乘得到的潜在 DOM 槽位数。 */
  projectedSlotCount: number;
};

/** 计算工作表性能画像所需的提前结构信息。 */
export type SpreadsheetPerformanceInput = {
  rowCount: number;
  columnCount: number;
  compressedBytes?: number;
  uncompressedBytes?: number;
  sheetBytes?: number;
  cfbFileBytes?: number;
  cfbMainStreamBytes?: number;
  modelBuildMilliseconds?: number;
};

/** 按边界值选择工作表模式；边界本身即进入优化路径。 */
export function createSpreadsheetPerformanceProfile({
  rowCount,
  columnCount,
  compressedBytes = 0,
  uncompressedBytes = 0,
  sheetBytes = 0,
  cfbFileBytes = 0,
  cfbMainStreamBytes = 0,
  modelBuildMilliseconds = 0,
}: SpreadsheetPerformanceInput): SpreadsheetPerformanceProfile {
  const safeRows = Math.max(1, Math.floor(rowCount || 1));
  const safeColumns = Math.max(1, Math.floor(columnCount || 1));
  const projectedSlotCount = safeRows * safeColumns;
  const gridMode =
    safeColumns >= SPREADSHEET_COLUMN_WINDOW_THRESHOLD
      ? 'two-axis-window'
      : safeRows >= SPREADSHEET_ROW_WINDOW_THRESHOLD ||
        projectedSlotCount >= SPREADSHEET_PROJECTED_SLOT_THRESHOLD
      ? 'row-window'
      : 'table';
  const sheetMode =
    compressedBytes >= OFFICE_LARGE_FILE_THRESHOLDS.ooxmlCompressedBytes ||
    uncompressedBytes >= OFFICE_LARGE_FILE_THRESHOLDS.ooxmlUncompressedBytes ||
    sheetBytes >= OFFICE_LARGE_FILE_THRESHOLDS.ooxmlMainXmlBytes ||
    cfbFileBytes >= OFFICE_LARGE_FILE_THRESHOLDS.cfbFileBytes ||
    cfbMainStreamBytes >= OFFICE_LARGE_FILE_THRESHOLDS.cfbMainStreamBytes ||
    modelBuildMilliseconds >= OFFICE_LARGE_FILE_THRESHOLDS.slowTaskMilliseconds
      ? 'lazy'
      : 'materialized';
  return { sheetMode, gridMode, projectedSlotCount };
}

/** 只允许同一会话的性能模式升级，避免视图在滚动期间退回完整表格。 */
export function upgradeSpreadsheetPerformanceProfile(
  current: SpreadsheetPerformanceProfile,
  next: SpreadsheetPerformanceProfile,
): SpreadsheetPerformanceProfile {
  const gridRank = { table: 0, 'row-window': 1, 'two-axis-window': 2 } as const;
  return {
    sheetMode:
      current.sheetMode === 'lazy' || next.sheetMode === 'lazy'
        ? 'lazy'
        : 'materialized',
    gridMode:
      gridRank[next.gridMode] > gridRank[current.gridMode]
        ? next.gridMode
        : current.gridMode,
    projectedSlotCount: Math.max(
      current.projectedSlotCount,
      next.projectedSlotCount,
    ),
  };
}
