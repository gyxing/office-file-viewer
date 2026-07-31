import { throwIfSpreadsheetAborted } from '../../spreadsheet/SpreadsheetSource';

/** BIFF8 SST 的按引用读取协议。 */
export interface Biff8SharedStringSource {
  /** 批量解析实际使用的共享字符串索引。 */
  resolveMany(
    indexes: readonly number[],
    signal?: AbortSignal,
  ): Promise<ReadonlyMap<number, string>>;
  /** 幂等释放当前对象持有的资源和订阅。 */
  dispose(): Promise<void>;
}

/** 为已解析的 Globals SST 提供统一按引用接口。 */
export function createBiff8SharedStringSource(
  sharedStrings: readonly string[],
): Biff8SharedStringSource {
  let disposed = false;
  return {
    async resolveMany(indexes, signal) {
      if (disposed) throw new Error('BIFF8 SST Source 已释放');
      throwIfSpreadsheetAborted(signal);
      return new Map(
        indexes.map((index) => [index, sharedStrings[index] ?? ''] as const),
      );
    },
    async dispose() {
      disposed = true;
    },
  };
}
