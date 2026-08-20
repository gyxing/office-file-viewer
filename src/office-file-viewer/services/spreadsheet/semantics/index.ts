export { applySpreadsheetTableSemantics } from './applySpreadsheetTableSemantics';
export { ConditionalFormattingIndex } from './ConditionalFormattingIndex';
export {
  applyConditionalFormatting,
  conditionalRuleNeedsFullRangeStats,
  createConditionalFormattingStatsAccumulator,
} from './evaluateConditionalFormatting';
export type { ConditionalFormattingStats } from './evaluateConditionalFormatting';
export type {
  SpreadsheetAnnotation,
  SpreadsheetAutoFilter,
  SpreadsheetConditionalFormattingRule,
  SpreadsheetConditionalValue,
  SpreadsheetPane,
  SpreadsheetTable,
} from './types';
