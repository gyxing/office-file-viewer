import type { SpreadsheetCellStyle, SpreadsheetRange } from '../types';

/** 工作表冻结或普通拆分窗格语义。 */
export type SpreadsheetPane = {
  /** 顶部冻结的行数。 */
  frozenRows: number;
  /** 左侧冻结的列数。 */
  frozenColumns: number;
  /** 滚动主体左上角使用的 A1 单元格。 */
  topLeftCell?: string;
  /** 源文件当前激活的窗格标识。 */
  activePane?: string;
  /** frozen 和 frozenSplit 可渲染；split 仅保留元数据。 */
  state?: 'frozen' | 'frozenSplit' | 'split';
  /** 普通拆分或 frozenSplit 的水平拆分位置。 */
  splitX?: number;
  /** 普通拆分或 frozenSplit 的垂直拆分位置。 */
  splitY?: number;
};

/** Excel Table 的样式和范围语义。 */
export type SpreadsheetTable = {
  /** 工作表内稳定的 Table 标识。 */
  id: string;
  /** 源文件声明的 Table 名称。 */
  name: string;
  /** Table 覆盖的 A1 范围。 */
  ref: string;
  /** 已解析的行列范围。 */
  range: SpreadsheetRange;
  /** 是否显示表头行。 */
  headerRow: boolean;
  /** 是否显示汇总行。 */
  totalsRow: boolean;
  /** 源 Table 样式名称。 */
  styleName?: string;
  /** 是否显示行条纹。 */
  showRowStripes?: boolean;
  /** 是否显示列条纹。 */
  showColumnStripes?: boolean;
  /** Table 内部的静态筛选范围。 */
  autoFilterRef?: string;
};

/** 工作表级 AutoFilter 及源文件保存的筛选状态。 */
export type SpreadsheetAutoFilter = {
  /** AutoFilter 覆盖的 A1 范围。 */
  ref: string;
  /** 已解析的行列范围。 */
  range?: SpreadsheetRange;
  /** 当前存在筛选条件的零基字段索引。 */
  filteredColumns?: readonly number[];
};

/** 单元格批注及共享审阅面板需要的作者和正文。 */
export type SpreadsheetAnnotation = {
  /** 工作表中稳定且唯一的批注标识。 */
  id: string;
  /** 批注锚定的 A1 单元格。 */
  ref: string;
  /** 从 1 开始的行索引。 */
  row: number;
  /** 从 1 开始的列索引。 */
  column: number;
  /** 批注作者；源文件缺失时保持为空。 */
  author?: string;
  /** 源文件提供的 ISO 日期或原始日期文本。 */
  createdAt?: string;
  /** 回复所属的父批注标识。 */
  parentId?: string;
  /** 批注线程是否已经解决。 */
  resolved?: boolean;
  /** 批注正文纯文本。 */
  text: string;
};

/** 条件格式阈值或公式操作数。 */
export type SpreadsheetConditionalValue = {
  /** 数字、百分比、公式或范围最值。 */
  type: string;
  /** 源文件声明的原始值。 */
  value?: string;
};

/** 首期可渲染和可保留摘要的条件格式规则。 */
export type SpreadsheetConditionalFormattingRule = {
  /** 当前工作表中的稳定规则标识。 */
  id: string;
  /** Excel 规则类型。 */
  type:
    | 'cellIs'
    | 'colorScale'
    | 'dataBar'
    | 'iconSet'
    | 'duplicateValues'
    | 'uniqueValues'
    | 'top10'
    | 'aboveAverage'
    | 'expression'
    | 'unsupported';
  /** 数值比较操作符。 */
  operator?: string;
  /** Excel 计算规则优先级，数值越小越先执行。 */
  priority: number;
  /** 当前规则命中后是否停止后续规则。 */
  stopIfTrue?: boolean;
  /** 规则覆盖的一个或多个 A1 范围。 */
  ranges: readonly SpreadsheetRange[];
  /** 条件格式引用的公式或阈值。 */
  values?: readonly SpreadsheetConditionalValue[];
  /** 条件格式命中时覆盖的差异样式。 */
  style?: SpreadsheetCellStyle;
  /** 色阶使用的 CSS 颜色。 */
  colors?: readonly string[];
  /** 数据条使用的 CSS 颜色。 */
  dataBarColor?: string;
  /** 图标集名称。 */
  iconSet?: string;
  /** top10 规则的排名或百分比。 */
  rank?: number;
  /** aboveAverage 是否匹配低于平均值。 */
  belowAverage?: boolean;
};
