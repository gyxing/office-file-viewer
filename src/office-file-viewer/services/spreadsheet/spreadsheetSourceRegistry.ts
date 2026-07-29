import { createMaterializedSpreadsheetSource } from './createMaterializedSpreadsheetSource';
import type { SpreadsheetSource } from './SpreadsheetSource';
import type { SpreadsheetWorkbook } from './types';

const sources = new WeakMap<SpreadsheetWorkbook, SpreadsheetSource>();

/** 为完整工作簿注册已存在的 Source。 */
export function registerSpreadsheetSource(
  workbook: SpreadsheetWorkbook,
  source: SpreadsheetSource,
) {
  sources.set(workbook, source);
}

/** 获取工作簿 Source；普通工作簿在首次读取时创建零拷贝适配器。 */
export function getSpreadsheetSource(workbook: SpreadsheetWorkbook) {
  let source = sources.get(workbook);
  if (!source) {
    source = createMaterializedSpreadsheetSource(workbook);
    sources.set(workbook, source);
  }
  return source;
}
