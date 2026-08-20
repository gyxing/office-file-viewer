import type { SpreadsheetCell, SpreadsheetRange } from '../types';
import type { SpreadsheetConditionalFormattingRule } from './types';

/** 单条规则覆盖范围的可复用统计。 */
export type ConditionalFormattingStats = {
  /** 规则范围内非空显示值的出现次数。 */
  valueCounts: Map<string, number>;
  /** 数字平均值。 */
  average?: number;
  /** 数字最小值。 */
  minimum?: number;
  /** 数字最大值。 */
  maximum?: number;
  /** top10 规则预先计算的排名阈值。 */
  topThreshold?: number;
};

function contains(range: SpreadsheetRange, cell: SpreadsheetCell) {
  return (
    cell.rowIndex >= range.startRow &&
    cell.rowIndex <= range.endRow &&
    cell.columnIndex >= range.startColumn &&
    cell.columnIndex <= range.endColumn
  );
}

function matchesRuleRange(
  rule: SpreadsheetConditionalFormattingRule,
  cell: SpreadsheetCell,
) {
  return rule.ranges.some((range) => contains(range, cell));
}

function numericCellValue(cell: SpreadsheetCell) {
  const value = Number(cell.rawValue ?? cell.value);
  return Number.isFinite(value) ? value : undefined;
}

/** 为需要全范围信息的规则计算一次统计并缓存。 */
/** 判断规则是否依赖完整目标范围，而不能只使用当前可见单元格。 */
export function conditionalRuleNeedsFullRangeStats(
  rule: SpreadsheetConditionalFormattingRule,
) {
  return !['cellIs', 'expression', 'unsupported'].includes(rule.type);
}

/** 创建可分片累加的条件格式统计，避免 Source 拼接整张表。 */
export function createConditionalFormattingStatsAccumulator(
  rule: SpreadsheetConditionalFormattingRule,
) {
  let count = 0;
  let sum = 0;
  let minimum: number | undefined;
  let maximum: number | undefined;
  const numbers = rule.type === 'top10' ? ([] as number[]) : undefined;
  const valueCounts = new Map<string, number>();
  return {
    add(cells: readonly SpreadsheetCell[]) {
      cells.forEach((cell) => {
        if (!matchesRuleRange(rule, cell)) return;
        const number = numericCellValue(cell);
        if (number !== undefined) {
          count += 1;
          sum += number;
          minimum = minimum === undefined ? number : Math.min(minimum, number);
          maximum = maximum === undefined ? number : Math.max(maximum, number);
          numbers?.push(number);
        }
        if (
          cell.value &&
          (rule.type === 'duplicateValues' || rule.type === 'uniqueValues')
        ) {
          valueCounts.set(cell.value, (valueCounts.get(cell.value) ?? 0) + 1);
        }
      });
    },
    finish(): ConditionalFormattingStats {
      const rank = Math.max(1, Math.trunc(rule.rank ?? 10));
      const sorted = numbers?.sort((left, right) => right - left);
      return {
        valueCounts,
        average: count ? sum / count : undefined,
        minimum,
        maximum,
        topThreshold: sorted?.length
          ? sorted[Math.min(sorted.length, rank) - 1]
          : undefined,
      };
    },
  };
}

/** 为一组已物化单元格计算规则统计。 */
export function buildConditionalFormattingStats(
  cells: readonly SpreadsheetCell[],
  rule: SpreadsheetConditionalFormattingRule,
): ConditionalFormattingStats {
  const accumulator = createConditionalFormattingStatsAccumulator(rule);
  accumulator.add(cells);
  return accumulator.finish();
}

function numericRuleValues(rule: SpreadsheetConditionalFormattingRule) {
  return (rule.values ?? []).flatMap((value) => {
    const number = Number(value.value);
    return Number.isFinite(number) ? [number] : [];
  });
}

function compareCellIs(
  value: number,
  rule: SpreadsheetConditionalFormattingRule,
) {
  const [first, second = first] = numericRuleValues(rule);
  if (first === undefined) return false;
  if (rule.operator === 'lessThan') return value < first;
  if (rule.operator === 'lessThanOrEqual') return value <= first;
  if (rule.operator === 'greaterThan') return value > first;
  if (rule.operator === 'greaterThanOrEqual') return value >= first;
  if (rule.operator === 'notEqual') return value !== first;
  if (rule.operator === 'between') return value >= first && value <= second;
  if (rule.operator === 'notBetween') return value < first || value > second;
  return value === first;
}

function interpolateHex(left: string, right: string, ratio: number) {
  const read = (value: string) => Number.parseInt(value.replace('#', ''), 16);
  const leftValue = read(left);
  const rightValue = read(right);
  const channel = (shift: number) =>
    Math.round(
      ((leftValue >> shift) & 255) * (1 - ratio) +
        ((rightValue >> shift) & 255) * ratio,
    );
  return `#${[channel(16), channel(8), channel(0)]
    .map((value) => value.toString(16).padStart(2, '0'))
    .join('')}`;
}

function colorScaleStyle(
  value: number,
  rule: SpreadsheetConditionalFormattingRule,
  stats: ConditionalFormattingStats,
) {
  if (
    stats.minimum === undefined ||
    stats.maximum === undefined ||
    !rule.colors?.length
  ) {
    return undefined;
  }
  const ratio =
    stats.maximum === stats.minimum
      ? 1
      : Math.max(
          0,
          Math.min(
            1,
            (value - stats.minimum) / (stats.maximum - stats.minimum),
          ),
        );
  if (rule.colors.length < 3) {
    return {
      backgroundColor: interpolateHex(
        rule.colors[0],
        rule.colors[rule.colors.length - 1],
        ratio,
      ),
    };
  }
  const lower = ratio <= 0.5;
  return {
    backgroundColor: interpolateHex(
      lower ? rule.colors[0] : rule.colors[1],
      lower ? rule.colors[1] : rule.colors[2],
      lower ? ratio * 2 : (ratio - 0.5) * 2,
    ),
  };
}

/** 判断规则是否命中，并返回样式或数据条/图标集视觉。 */
function evaluateRule(
  cell: SpreadsheetCell,
  rule: SpreadsheetConditionalFormattingRule,
  stats: ConditionalFormattingStats,
) {
  const value = numericCellValue(cell);
  if (rule.type === 'cellIs') {
    return value !== undefined && compareCellIs(value, rule)
      ? { style: rule.style }
      : undefined;
  }
  if (rule.type === 'duplicateValues' || rule.type === 'uniqueValues') {
    const count = stats.valueCounts.get(cell.value) ?? 0;
    const matched = rule.type === 'duplicateValues' ? count > 1 : count === 1;
    return matched ? { style: rule.style } : undefined;
  }
  if (rule.type === 'aboveAverage') {
    if (value === undefined || stats.average === undefined) return undefined;
    const matched = rule.belowAverage
      ? value < stats.average
      : value > stats.average;
    return matched ? { style: rule.style } : undefined;
  }
  if (rule.type === 'top10') {
    if (value === undefined || stats.topThreshold === undefined) {
      return undefined;
    }
    return value >= stats.topThreshold ? { style: rule.style } : undefined;
  }
  if (rule.type === 'colorScale' && value !== undefined) {
    const style = colorScaleStyle(value, rule, stats);
    return style ? { style } : undefined;
  }
  if (rule.type === 'dataBar' && value !== undefined) {
    if (stats.minimum === undefined || stats.maximum === undefined)
      return undefined;
    const percent =
      stats.maximum === stats.minimum
        ? 100
        : ((value - stats.minimum) / (stats.maximum - stats.minimum)) * 100;
    return {
      visual: {
        dataBarPercent: Math.max(0, Math.min(100, percent)),
        dataBarColor: rule.dataBarColor ?? '#5b9bd5',
      },
    };
  }
  if (rule.type === 'iconSet' && value !== undefined) {
    if (stats.minimum === undefined || stats.maximum === undefined)
      return undefined;
    const ratio =
      stats.maximum === stats.minimum
        ? 1
        : (value - stats.minimum) / (stats.maximum - stats.minimum);
    return {
      visual: {
        iconIndex: ratio >= 2 / 3 ? 2 : ratio >= 1 / 3 ? 1 : 0,
        iconSet: rule.iconSet ?? '3TrafficLights1',
      },
    };
  }
  return undefined;
}

/** 对完整或当前 Source 范围的单元格应用可安全求值的条件格式。 */
export function applyConditionalFormatting(
  cells: SpreadsheetCell[],
  rules: readonly SpreadsheetConditionalFormattingRule[],
  precomputedStats?: ReadonlyMap<string, ConditionalFormattingStats>,
) {
  const stats = new Map(
    rules.map(
      (rule) =>
        [
          rule.id,
          precomputedStats?.get(rule.id) ??
            buildConditionalFormattingStats(cells, rule),
        ] as const,
    ),
  );
  cells.forEach((cell) => {
    for (const rule of [...rules].sort((a, b) => a.priority - b.priority)) {
      if (!matchesRuleRange(rule, cell)) continue;
      const result = evaluateRule(cell, rule, stats.get(rule.id)!);
      if (!result) continue;
      if (result.style) cell.style = { ...cell.style, ...result.style };
      if (result.visual) cell.conditionalVisual = result.visual;
      if (rule.stopIfTrue) break;
    }
  });
}
