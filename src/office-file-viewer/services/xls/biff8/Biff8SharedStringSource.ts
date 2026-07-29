import { throwIfSpreadsheetAborted } from '../../spreadsheet/SpreadsheetSource';

/** BIFF8 SST 的按引用读取协议。 */
export interface Biff8SharedStringSource {
  resolveMany(
    indexes: readonly number[],
    signal?: AbortSignal,
  ): Promise<ReadonlyMap<number, string>>;
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
