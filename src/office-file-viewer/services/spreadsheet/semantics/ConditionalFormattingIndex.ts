import type { SpreadsheetRange } from '../types';
import type { SpreadsheetConditionalFormattingRule } from './types';

/** 判断两个闭合行列范围是否相交。 */
function intersects(left: SpreadsheetRange, right: SpreadsheetRange) {
  return !(
    left.endRow < right.startRow ||
    left.startRow > right.endRow ||
    left.endColumn < right.startColumn ||
    left.startColumn > right.endColumn
  );
}

/** 按范围查询条件格式规则，避免渲染器逐单元格扫描整张规则表。 */
export class ConditionalFormattingIndex {
  private readonly entries: Array<{
    range: SpreadsheetRange;
    rule: SpreadsheetConditionalFormattingRule;
  }>;

  constructor(rules: readonly SpreadsheetConditionalFormattingRule[]) {
    this.entries = rules
      .flatMap((rule) => rule.ranges.map((range) => ({ range, rule })))
      .sort(
        (left, right) =>
          left.range.startRow - right.range.startRow ||
          left.range.startColumn - right.range.startColumn ||
          left.rule.priority - right.rule.priority,
      );
  }

  /** 返回与当前可见范围相交且按优先级排序的规则。 */
  query(range: SpreadsheetRange) {
    const rules = new Map<string, SpreadsheetConditionalFormattingRule>();
    this.entries.forEach((entry) => {
      if (entry.range.startRow > range.endRow) return;
      if (intersects(entry.range, range)) rules.set(entry.rule.id, entry.rule);
    });
    return [...rules.values()].sort(
      (left, right) => left.priority - right.priority,
    );
  }
}
