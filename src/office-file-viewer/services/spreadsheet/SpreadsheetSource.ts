import type { OfficeSearchProvider } from '../search/types';
import type { SpreadsheetPerformanceProfile } from './spreadsheetPerformance';
import type {
  SpreadsheetColumnMetric,
  SpreadsheetRange,
  SpreadsheetRangeData,
  SpreadsheetRowMetric,
  SpreadsheetSheet,
} from './types';

/** 描述按需电子表格中的轻量 Sheet 状态。 */
export type SpreadsheetSheetDescriptor = {
  /** Sheet 稳定标识。 */
  id: string;
  /** Sheet 展示名称。 */
  name: string;
  /** Sheet 在源容器中的路径或子流位置。 */
  path: string;
  /** 普通工作表或图表工作表。 */
  kind: 'worksheet' | 'chart';
  /** 当前估算或确认的总行数。 */
  rowCount: number;
  /** 当前估算或确认的总列数。 */
  columnCount: number;
  /** 描述符版本。 */
  revision: number;
  /** 尚未解析、已就绪或解析失败。 */
  status: 'estimated' | 'ready' | 'error';
  /** 解析失败时的诊断文本。 */
  errorMessage?: string;
};

/** React 订阅使用的不可变工作簿快照。 */
export type SpreadsheetSourceSnapshot = {
  /** 快照版本。 */
  revision: number;
  /** 按源工作簿顺序排列的 Sheet 描述符。 */
  sheets: readonly SpreadsheetSheetDescriptor[];
  /** 工作簿级定义名称到静态目标地址的映射。 */
  definedNames?: Readonly<Record<string, string>>;
};

/** 虚拟网格计算全局坐标所需的轻量 Sheet 布局。 */
export type SpreadsheetSheetLayout = {
  /** 布局版本。 */
  revision: number;
  /** 工作表总行数。 */
  rowCount: number;
  /** 工作表总列数。 */
  columnCount: number;
  /** 未显式覆盖时使用的默认行高。 */
  defaultRowHeight: number;
  /** 未显式覆盖时使用的默认列宽。 */
  defaultColumnWidth: number;
  /** 稀疏行尺寸覆盖。 */
  rows: readonly SpreadsheetRowMetric[];
  /** 稀疏列尺寸覆盖。 */
  columns: readonly SpreadsheetColumnMetric[];
};

/** 为普通和大型工作簿提供统一的 Sheet 与范围读取协议。 */
export interface SpreadsheetSource {
  /** 当前工作簿支持搜索时提供的稀疏扫描能力。 */
  readonly searchProvider?: OfficeSearchProvider;
  /** 返回当前可观察状态的只读快照。 */
  getSnapshot(): SpreadsheetSourceSnapshot;
  /** 订阅状态快照变化，并返回取消订阅函数。 */
  subscribe(listener: () => void): () => void;
  /** 返回指定工作表采用的性能配置。 */
  getProfile(sheetId: string): SpreadsheetPerformanceProfile;
  /** 读取指定工作表的行列布局。 */
  getSheetLayout(sheetId: string): SpreadsheetSheetLayout;
  /** 确保指定工作表的结构数据已经加载。 */
  ensureSheet(sheetId: string, signal?: AbortSignal): Promise<void>;
  /** 返回已完整物化的工作表；不可用时返回空值。 */
  getMaterializedSheet(
    sheetId: string,
    signal?: AbortSignal,
  ): Promise<SpreadsheetSheet | undefined>;
  /** 读取指定工作表范围内的单元格与对象。 */
  getRange(
    sheetId: string,
    range: SpreadsheetRange,
    signal?: AbortSignal,
  ): Promise<SpreadsheetRangeData>;
  /** 保留指定可视范围并回收远离窗口的缓存内容。 */
  retainRange(sheetId: string, range: SpreadsheetRange): () => void;
  /** 重新加载此前失败的工作表。 */
  retrySheet(sheetId: string): void;
  /** 幂等释放当前对象持有的资源和订阅。 */
  dispose(): Promise<void>;
}

/** 创建与浏览器 AbortSignal 一致的电子表格取消错误。 */
export function createSpreadsheetAbortError(message = '工作表读取已取消') {
  const error = new Error(message);
  error.name = 'AbortError';
  return error;
}

/** 在进入可取消的工作表步骤前统一检查取消状态。 */
export function throwIfSpreadsheetAborted(signal?: AbortSignal) {
  if (signal?.aborted) throw createSpreadsheetAbortError();
}

/** 让多个调用者共享解析任务，同时保持每个调用者可独立取消等待。 */
export function waitForSpreadsheetResult<T>(
  promise: Promise<T>,
  signal?: AbortSignal,
): Promise<T> {
  if (!signal) return promise;
  if (signal.aborted) return Promise.reject(createSpreadsheetAbortError());
  return new Promise<T>((resolve, reject) => {
    const abort = () => reject(createSpreadsheetAbortError());
    signal.addEventListener('abort', abort, { once: true });
    promise.then(
      (value) => {
        signal.removeEventListener('abort', abort);
        resolve(value);
      },
      (error) => {
        signal.removeEventListener('abort', abort);
        reject(error);
      },
    );
  });
}
