import type { SpreadsheetMerge, SpreadsheetRange } from './types';

/** 判断两个包含结束位置的工作表范围是否相交。 */
export function spreadsheetRangesIntersect(
  left: SpreadsheetRange,
  right: SpreadsheetRange,
) {
  return !(
    left.endRow < right.startRow ||
    left.startRow > right.endRow ||
    left.endColumn < right.startColumn ||
    left.startColumn > right.endColumn
  );
}

/** 按行区间查询合并区域，并把虚拟窗口扩展到完整合并范围。 */
export class SpreadsheetMergeIndex {
  private readonly merges: readonly SpreadsheetMerge[];

  constructor(merges: readonly SpreadsheetMerge[]) {
    this.merges = [...merges].sort(
      (left, right) => left.startRow - right.startRow,
    );
  }

  query(range: SpreadsheetRange) {
    return this.merges.filter(
      (merge) =>
        merge.startRow <= range.endRow &&
        spreadsheetRangesIntersect(range, merge),
    );
  }

  expand(range: SpreadsheetRange) {
    let expanded = { ...range };
    let changed = true;
    while (changed) {
      changed = false;
      for (const merge of this.query(expanded)) {
        const next = {
          startRow: Math.min(expanded.startRow, merge.startRow),
          endRow: Math.max(expanded.endRow, merge.endRow),
          startColumn: Math.min(expanded.startColumn, merge.startColumn),
          endColumn: Math.max(expanded.endColumn, merge.endColumn),
        };
        if (
          next.startRow !== expanded.startRow ||
          next.endRow !== expanded.endRow ||
          next.startColumn !== expanded.startColumn ||
          next.endColumn !== expanded.endColumn
        ) {
          expanded = next;
          changed = true;
        }
      }
    }
    return expanded;
  }
}
