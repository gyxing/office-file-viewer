// PPT 文本解析入口，统一导出文本原子、样式和标准化文本结果。
export { mergePptTextStyles } from './mergeTextStyles';
export { parsePptText, parsePptTextGroups } from './parsePptText';
export { readPptTextAtoms } from './readTextAtoms';
export { readPptTextStyles } from './readTextStyles';
export type {
  PptCharacterStyleRun,
  PptParagraphStyleRun,
  PptParsedText,
  PptTextAtomGroup,
  PptTextDefaults,
  PptTextStyleRuns,
} from './types';
