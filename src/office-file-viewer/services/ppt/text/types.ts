import type { TextParagraph, TextStyle } from '../../presentation/types';
import type { PptRecord } from '../types';

/** PPT 文档、母版和占位符层级的文字默认样式。 */
export type PptTextDefaults = {
  /** 文档级默认文字样式。 */
  document?: TextStyle;
  /** 母版级默认文字样式。 */
  master?: TextStyle;
  /** 占位符级默认文字样式。 */
  placeholder?: TextStyle;
  /** 按字体编号索引的字体族名称。 */
  fonts?: Map<number, string>;
};

/** PPT 文本内容记录及其配套样式记录。 */
export type PptTextAtomGroup = {
  /** PPT TextHeaderAtom 声明的文本用途编号。 */
  textType: number;
  /** 文本内容。 */
  text: string;
  /** 保存实际文本内容的 PPT 记录。 */
  contentRecord: PptRecord;
  /** 保存文本样式区间的 PPT 记录。 */
  styleRecord?: PptRecord;
};

/** PPT 段落样式覆盖的字符数量和层级。 */
export type PptParagraphStyleRun = {
  /** 当前集合或范围包含的项目数量。 */
  count: number;
  /** 编号或大纲的零基级别。 */
  level: number;
  /** 当前内容使用的渲染样式。 */
  style: TextStyle;
};

/** PPT 字符样式覆盖的字符数量。 */
export type PptCharacterStyleRun = {
  /** 当前集合或范围包含的项目数量。 */
  count: number;
  /** 当前内容使用的渲染样式。 */
  style: TextStyle;
};

/** PPT 段落和字符样式区间集合。 */
export type PptTextStyleRuns = {
  /** 按源文档顺序排列的段落。 */
  paragraphs: PptParagraphStyleRun[];
  /** 按字符范围排列的文字样式区间。 */
  characters: PptCharacterStyleRun[];
};

/** 完成样式合并的 PPT 文本段落。 */
export type PptParsedText = {
  /** PPT TextHeaderAtom 声明的文本用途编号。 */
  textType: number;
  /** 按源文档顺序排列的段落。 */
  paragraphs: TextParagraph[];
};
