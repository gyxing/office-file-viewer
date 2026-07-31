import type { SpreadsheetColumnMetric, SpreadsheetRowMetric } from './types';

/** 提供稀疏行列尺寸的全局像素坐标查询。 */
export interface SpreadsheetAxisIndex {
  /** 计算指定行或列起点的累计偏移。 */
  offsetAt(index: number): number;
  /** 返回指定行或列的显示尺寸。 */
  sizeAt(index: number): number;
  /** 计算连续行列范围的累计尺寸。 */
  rangeSize(start: number, end: number): number;
  /** 查找指定轴向偏移落入的行或列索引。 */
  findIndexAtOffset(offset: number): number;
}

/** 电子表格单个行轴或列轴的尺寸信息。 */
type AxisMetric = SpreadsheetRowMetric | SpreadsheetColumnMetric;

/** 创建不会按 Excel 最大行列数分配等长数组的轴索引。 */
export function createSpreadsheetAxisIndex(
  count: number,
  defaultSize: number,
  metrics: readonly AxisMetric[],
): SpreadsheetAxisIndex {
  const safeCount = Math.max(1, Math.floor(count || 1));
  const safeDefault = Math.max(0, Number(defaultSize) || 0);
  const overrides = [...metrics]
    .filter(
      (metric) =>
        Number.isInteger(metric.index) &&
        metric.index >= 1 &&
        metric.index <= safeCount,
    )
    .sort((left, right) => left.index - right.index)
    .map((metric) => ({
      index: metric.index,
      size: metric.hidden
        ? 0
        : Math.max(
            0,
            Number('height' in metric ? metric.height : metric.width),
          ),
    }));
  const byIndex = new Map(overrides.map((item) => [item.index, item.size]));
  const prefixDeltas: number[] = [];
  let delta = 0;
  overrides.forEach((item) => {
    delta += item.size - safeDefault;
    prefixDeltas.push(delta);
  });

  const deltaBefore = (index: number) => {
    let low = 0;
    let high = overrides.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (overrides[middle].index < index) low = middle + 1;
      else high = middle;
    }
    return low ? prefixDeltas[low - 1] : 0;
  };
  const offsetAt = (index: number) => {
    const normalized = Math.max(1, Math.min(safeCount + 1, Math.floor(index)));
    return (normalized - 1) * safeDefault + deltaBefore(normalized);
  };
  const sizeAt = (index: number) => {
    if (index < 1 || index > safeCount) return 0;
    return byIndex.get(Math.floor(index)) ?? safeDefault;
  };

  return {
    offsetAt,
    sizeAt,
    rangeSize(start, end) {
      if (end < start) return 0;
      return offsetAt(Math.min(safeCount + 1, end + 1)) - offsetAt(start);
    },
    findIndexAtOffset(offset) {
      const target = Math.max(0, Number(offset) || 0);
      let low = 1;
      let high = safeCount;
      while (low < high) {
        const middle = Math.floor((low + high) / 2);
        if (offsetAt(middle + 1) <= target) low = middle + 1;
        else high = middle;
      }
      return low;
    },
  };
}
