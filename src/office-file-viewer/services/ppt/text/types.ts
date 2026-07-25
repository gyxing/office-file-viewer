import type { TextParagraph, TextStyle } from '../../presentation/types';
import type { PptRecord } from '../types';

/** 描述 PptTextDefaults 在 PPT 二进制解析中的数据结构。 */
export type PptTextDefaults = {
  /** PptTextDefaults 当前关联的标准化文档模型。 */
  document?: TextStyle;
  /** PptTextDefaults 关联的 master 结构；字段形状由 TextStyle 定义；未提供时使用来源格式或渲染器的默认行为。 */
  master?: TextStyle;
  /** PptTextDefaults 关联的 placeholder 结构；字段形状由 TextStyle 定义；未提供时使用来源格式或渲染器的默认行为。 */
  placeholder?: TextStyle;
  /** PptTextDefaults 按业务键索引的 fonts 映射；未提供时使用来源格式或渲染器的默认行为。 */
  fonts?: Map<number, string>;
};

/** 描述 PptTextAtomGroup 在 PPT 二进制解析中的数据结构。 */
export type PptTextAtomGroup = {
  /** PptTextAtomGroup 从源格式读取的 textType 枚举或标识值。 */
  textType: number;
  /** PptTextAtomGroup 携带或渲染的文本内容。 */
  text: string;
  /** PptTextAtomGroup 关联的 contentRecord 结构；字段形状由 PptRecord 定义。 */
  contentRecord: PptRecord;
  /** PptTextAtomGroup 关联的 styleRecord 结构；字段形状由 PptRecord 定义；未提供时使用来源格式或渲染器的默认行为。 */
  styleRecord?: PptRecord;
};

/** 描述 PptParagraphStyleRun 在 PPT 二进制解析中的数据结构。 */
export type PptParagraphStyleRun = {
  /** PptParagraphStyleRun 对应项目的数量。 */
  count: number;
  /** PptParagraphStyleRun 从源格式读取的 level 枚举或标识值。 */
  level: number;
  /** PptParagraphStyleRun 使用的渲染或文本样式。 */
  style: TextStyle;
};

/** 描述 PptCharacterStyleRun 在 PPT 二进制解析中的数据结构。 */
export type PptCharacterStyleRun = {
  /** PptCharacterStyleRun 对应项目的数量。 */
  count: number;
  /** PptCharacterStyleRun 使用的渲染或文本样式。 */
  style: TextStyle;
};

/** 描述 PptTextStyleRuns 在 PPT 二进制解析中的数据结构。 */
export type PptTextStyleRuns = {
  /** PptTextStyleRuns 包含的 paragraphs 有序集合。 */
  paragraphs: PptParagraphStyleRun[];
  /** PptTextStyleRuns 包含的 characters 有序集合。 */
  characters: PptCharacterStyleRun[];
};

/** 描述 PptParsedText 在 PPT 二进制解析中的数据结构。 */
export type PptParsedText = {
  /** PptParsedText 从源格式读取的 textType 枚举或标识值。 */
  textType: number;
  /** PptParsedText 包含的 paragraphs 有序集合。 */
  paragraphs: TextParagraph[];
};
