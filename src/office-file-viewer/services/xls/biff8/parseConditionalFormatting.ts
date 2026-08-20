import type { SpreadsheetConditionalFormattingRule } from '../../spreadsheet/semantics/types';
import type {
  SpreadsheetRange,
  SpreadsheetWarning,
} from '../../spreadsheet/types';
import type { Biff8SheetDescriptor } from '../types';

/** 解析 BIFF8 CondFmt 的 SqRefU 范围集合。 */
export function parseBiff8ConditionalRanges(
  bytes: Uint8Array,
  descriptor: Biff8SheetDescriptor,
  warnings: SpreadsheetWarning[],
) {
  if (bytes.length < 14) return [];
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const count = view.getUint16(12, true);
  const ranges: SpreadsheetRange[] = [];
  for (let index = 0; index < count; index += 1) {
    const offset = 14 + index * 8;
    if (offset + 8 > bytes.length) {
      warnings.push({
        code: 'INVALID_CONDITIONAL_FORMATTING_RANGE',
        message: 'CondFmt 范围列表越界，已保留可恢复规则摘要',
        sheetName: descriptor.name,
      });
      break;
    }
    ranges.push({
      startRow: view.getUint16(offset, true) + 1,
      endRow: view.getUint16(offset + 2, true) + 1,
      startColumn: view.getUint16(offset + 4, true) + 1,
      endColumn: view.getUint16(offset + 6, true) + 1,
    });
  }
  return ranges;
}

/** 将 BIFF8 CF 比较函数映射为共享 cellIs 操作符。 */
function comparisonOperator(value: number) {
  return [
    undefined,
    'between',
    'notBetween',
    'equal',
    'notEqual',
    'greaterThan',
    'lessThan',
    'greaterThanOrEqual',
    'lessThanOrEqual',
  ][value];
}

/** 解析 BIFF8 CF 记录的类型与操作符，并保留公式摘要。 */
export function parseBiff8ConditionalRule(
  bytes: Uint8Array,
  ranges: readonly SpreadsheetRange[],
  priority: number,
  descriptor: Biff8SheetDescriptor,
  warnings: SpreadsheetWarning[],
): SpreadsheetConditionalFormattingRule | undefined {
  if (bytes.length < 6 || !ranges.length) return undefined;
  const conditionType = bytes[0] ?? 0;
  const comparison = bytes[1] ?? 0;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const formula1Length = view.getUint16(2, true);
  const formula2Length = view.getUint16(4, true);
  if (formula1Length || formula2Length) {
    warnings.push({
      code: 'XLS_CONDITIONAL_FORMAT_FORMULA_UNSUPPORTED',
      message: 'BIFF8 条件格式公式已保留规则摘要，当前不执行完整公式重算。',
      sheetName: descriptor.name,
    });
  }
  return {
    id: `${descriptor.id}:cf:${priority}`,
    type: conditionType === 1 ? 'cellIs' : 'expression',
    operator: comparisonOperator(comparison),
    priority,
    ranges: [...ranges],
    values: [
      ...(formula1Length ? [{ type: 'formula', value: 'BIFF8 formula' }] : []),
      ...(formula2Length ? [{ type: 'formula', value: 'BIFF8 formula' }] : []),
    ],
  };
}
