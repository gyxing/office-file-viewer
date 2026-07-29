import type { OfficeChartModel } from '../../shared/ooxml/charts';
import type { OfficeResourceSource } from '../resource-store';

/** 描述电子表格标准模型过程中可继续处理的警告。 */
export type SpreadsheetWarning = {
  /** SpreadsheetWarning 的稳定代码，用于程序化识别具体情况。 */
  code: string;
  /** SpreadsheetWarning 面向调用方或用户展示的具体警告、错误说明。 */
  message: string;
  /** SpreadsheetWarning 的 sheetName 文本值。 */
  sheetName?: string;
  /** SpreadsheetWarning 在所属数据范围中的偏移位置。 */
  offset?: number;
};

/** 描述 SpreadsheetWorkbook 在电子表格标准模型中的数据结构。 */
export type SpreadsheetWorkbook = {
  /** SpreadsheetWorkbook 包含的 sheets 有序集合。 */
  sheets: SpreadsheetSheet[];
  /** 解析时产生但不阻止继续预览的警告；未提供表示没有警告。 */
  warnings?: SpreadsheetWarning[];
  /** 工作簿持有且需要在销毁时释放的浏览器资源。 */
  resources?: SpreadsheetResources;
};

/** 记录电子表格标准模型持有且需要统一管理的资源。 */
export type SpreadsheetResources = {
  /** 浏览器创建的对象 URL 集合，文档释放时必须逐一撤销。 */
  objectUrls: string[];
};

/** 释放工作簿创建的 Blob URL；重复调用保持幂等。 */
export function disposeSpreadsheetWorkbook(
  workbook: SpreadsheetWorkbook | undefined,
) {
  const urls = workbook?.resources?.objectUrls;
  if (!urls?.length) return;
  const uniqueUrls = new Set(urls);
  urls.length = 0;
  if (typeof URL === 'undefined' || typeof URL.revokeObjectURL !== 'function') {
    return;
  }
  uniqueUrls.forEach((url) => URL.revokeObjectURL(url));
}

/** 描述 SpreadsheetSheet 在电子表格标准模型中的数据结构。 */
export type SpreadsheetSheet = {
  /** SpreadsheetSheet 在所属文档或任务中的唯一标识。 */
  id: string;
  /** SpreadsheetSheet 的可读名称。 */
  name: string;
  /** SpreadsheetSheet 在压缩包、复合文档或图形数据中的路径。 */
  path: string;
  /** 标识 SpreadsheetSheet 对应的 Office 文件或数据种类。 */
  kind?: 'worksheet' | 'chart';
  /** 工作表未单独设置列宽时使用的默认宽度，单位为标准化渲染像素。 */
  defaultColumnWidth?: number;
  /** 工作表未单独设置行高时使用的默认高度，单位为标准化渲染像素。 */
  defaultRowHeight?: number;
  /** 工作表有效内容范围，例如 A1:D20。 */
  range?: string;
  /** 工作表标准化后的总行数或总列数，具体含义由属性名确定。 */
  rowCount: number;
  /** 工作表标准化后的总行数或总列数，具体含义由属性名确定。 */
  columnCount: number;
  /** SpreadsheetSheet 包含的 columns 有序集合。 */
  columns: SpreadsheetColumn[];
  /** SpreadsheetSheet 包含的 rows 有序集合。 */
  rows: SpreadsheetRow[];
  /** SpreadsheetSheet 包含的 merges 有序集合。 */
  merges: SpreadsheetMerge[];
  /** SpreadsheetSheet 包含的 images 有序集合。 */
  images: SpreadsheetImage[];
  /** SpreadsheetSheet 包含的 charts 有序集合。 */
  charts: SpreadsheetChart[];
};

/** 描述 SpreadsheetColumn 在电子表格标准模型中的数据结构。 */
export type SpreadsheetColumn = {
  /** SpreadsheetColumn 在所属集合中的位置索引。 */
  index: number;
  /** SpreadsheetColumn 面向用户展示的标签文本。 */
  label: string;
  /** SpreadsheetColumn 的 width 尺寸或坐标，单位为标准化渲染像素。 */
  width: number;
  /** 是否隐藏 SpreadsheetColumn；未提供时沿用来源格式或渲染器的默认规则。 */
  hidden?: boolean;
};

/** 描述 SpreadsheetRow 在电子表格标准模型中的数据结构。 */
export type SpreadsheetRow = {
  /** SpreadsheetRow 在所属集合中的位置索引。 */
  index: number;
  /** SpreadsheetRow 的 height 尺寸或坐标，单位为标准化渲染像素。 */
  height: number;
  /** 是否隐藏 SpreadsheetRow；未提供时沿用来源格式或渲染器的默认规则。 */
  hidden?: boolean;
  /** SpreadsheetRow 包含的 cells 有序集合。 */
  cells: SpreadsheetCell[];
};

/** 描述 SpreadsheetCell 在电子表格标准模型中的数据结构。 */
export type SpreadsheetCell = {
  /** 单元格 A1 引用，例如 C5。 */
  ref: string;
  /** 单元格或记录所在的行索引。 */
  rowIndex: number;
  /** 单元格或记录所在的列索引。 */
  columnIndex: number;
  /** SpreadsheetCell 保存的解析值或业务值。 */
  value: string;
  /** 格式化前从源文件读取的原始值。 */
  rawValue?: string;
  /** 源文件中的单元格值类型标识。 */
  type?: string;
  /** 单元格引用的样式表索引。 */
  styleId?: number;
  /** SpreadsheetCell 使用的渲染或文本样式。 */
  style?: SpreadsheetCellStyle;
  /** 单元格公式文本，不包含计算结果。 */
  formula?: string;
  /** 二进制公式令牌无法完整还原时保留的可读表示。 */
  formulaTokens?: string;
  /** 表格单元格横向跨越的列数。 */
  colSpan?: number;
  /** 表格单元格纵向跨越的行数。 */
  rowSpan?: number;
  /** 是否因合并区域而隐藏该非主单元格。 */
  hiddenByMerge?: boolean;
};

/** 描述 SpreadsheetMerge 在电子表格标准模型中的数据结构。 */
export type SpreadsheetMerge = {
  /** SpreadsheetMerge 的 ref 文本值。 */
  ref: string;
  /** 合并区域起始行索引。 */
  startRow: number;
  /** 合并区域起始列索引。 */
  startColumn: number;
  /** 合并区域结束行索引。 */
  endRow: number;
  /** 合并区域结束列索引。 */
  endColumn: number;
};

/** 描述 SpreadsheetAnchorPoint 在电子表格标准模型中的数据结构。 */
export type SpreadsheetAnchorPoint = {
  /** 锚点所在的行索引。 */
  row: number;
  /** 锚点所在的列索引。 */
  column: number;
  /** 锚点在目标行内的垂直偏移，单位为标准化渲染像素。 */
  rowOffset: number;
  /** 锚点在目标列内的水平偏移，单位为标准化渲染像素。 */
  columnOffset: number;
};

/** 描述 SpreadsheetImage 在电子表格标准模型中的数据结构。 */
export type SpreadsheetImage = {
  /** SpreadsheetImage 在所属文档或任务中的唯一标识。 */
  id: string;
  /** SpreadsheetImage 的可读名称。 */
  name?: string;
  /** 图片可直接用于 img 元素的资源地址。 */
  src: string | OfficeResourceSource;
  /** 图片无法显示时使用的替代文本。 */
  alt?: string;
  /** SpreadsheetImage 的起始锚点。 */
  from: SpreadsheetAnchorPoint;
  /** SpreadsheetImage 的结束锚点。 */
  to: SpreadsheetAnchorPoint;
  /** SpreadsheetImage 的 x 尺寸或坐标，单位为标准化渲染像素。 */
  x: number;
  /** SpreadsheetImage 的 y 尺寸或坐标，单位为标准化渲染像素。 */
  y: number;
  /** SpreadsheetImage 的 width 尺寸或坐标，单位为标准化渲染像素。 */
  width: number;
  /** SpreadsheetImage 的 height 尺寸或坐标，单位为标准化渲染像素。 */
  height: number;
};

/** 描述 SpreadsheetChart 在电子表格标准模型中的数据结构。 */
export type SpreadsheetChart = {
  /** SpreadsheetChart 在所属文档或任务中的唯一标识。 */
  id: string;
  /** SpreadsheetChart 对外展示的标题。 */
  title?: string;
  /** SpreadsheetChart 当前关联的图表模型。 */
  chart: OfficeChartModel;
  /** SpreadsheetChart 的起始锚点。 */
  from: SpreadsheetAnchorPoint;
  /** SpreadsheetChart 的结束锚点。 */
  to: SpreadsheetAnchorPoint;
  /** SpreadsheetChart 的 x 尺寸或坐标，单位为标准化渲染像素。 */
  x: number;
  /** SpreadsheetChart 的 y 尺寸或坐标，单位为标准化渲染像素。 */
  y: number;
  /** SpreadsheetChart 的 width 尺寸或坐标，单位为标准化渲染像素。 */
  width: number;
  /** SpreadsheetChart 的 height 尺寸或坐标，单位为标准化渲染像素。 */
  height: number;
};

/** 描述单元格从源文件读取的对角边框；方向始终相对于完整单元格矩形。 */
export type SpreadsheetDiagonalBorder = {
  /** 是否绘制左下到右上的对角线。 */
  up: boolean;
  /** 是否绘制左上到右下的对角线。 */
  down: boolean;
  /** 对角边框颜色，使用标准化 CSS 颜色值。 */
  color: string;
  /** 对角边框宽度，单位为标准化渲染像素。 */
  width: number;
  /** 对角边框的源文件线型。 */
  lineStyle:
    | 'hair'
    | 'thin'
    | 'medium'
    | 'thick'
    | 'double'
    | 'dotted'
    | 'dashed'
    | 'dashDot'
    | 'dashDotDot'
    | 'slantDashDot';
};

/** 描述电子表格标准模型使用的样式参数。 */
export type SpreadsheetCellStyle = {
  /** 是否使用粗体渲染 SpreadsheetCellStyle；未提供时沿用来源格式或渲染器的默认规则。 */
  bold?: boolean;
  /** 是否使用斜体渲染 SpreadsheetCellStyle；未提供时沿用来源格式或渲染器的默认规则。 */
  italic?: boolean;
  /** 是否为 SpreadsheetCellStyle 绘制下划线；未提供时沿用来源格式或渲染器的默认规则。 */
  underline?: boolean;
  /** SpreadsheetCellStyle 的前景或文本颜色，使用标准化 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
  color?: string;
  /** SpreadsheetCellStyle 的字体族名称；未提供时沿用来源格式或渲染器的默认规则。 */
  fontFamily?: string;
  /** SpreadsheetCellStyle 的字号，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  fontSize?: number;
  /** SpreadsheetCellStyle 的背景颜色，使用 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
  backgroundColor?: string;
  /** 单元格内容的水平对齐方式。 */
  horizontalAlign?: 'left' | 'center' | 'right' | 'justify';
  /** SpreadsheetCellStyle 的垂直对齐方式；未提供时沿用来源格式或渲染器的默认规则。 */
  verticalAlign?: 'top' | 'middle' | 'bottom';
  /** 是否允许单元格文本自动换行。 */
  wrapText?: boolean;
  /** 是否存在任意方向的有效边框。 */
  border?: boolean;
  /** SpreadsheetCellStyle 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderTop?: string;
  /** SpreadsheetCellStyle 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderRight?: string;
  /** SpreadsheetCellStyle 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderBottom?: string;
  /** SpreadsheetCellStyle 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderLeft?: string;
  /** 未单独指定方向时使用的统一边框颜色。 */
  borderColor?: string;
  /** 未单独指定方向时使用的统一边框宽度，单位为像素。 */
  borderWidth?: number;
  /**
   * 单元格对角边框。
   *
   * 对角线属于单元格样式，其端点由单元格或合并单元格的最终边框盒决定，
   * 不能按工作表坐标猜测为浮动图形。
   */
  diagonalBorder?: SpreadsheetDiagonalBorder;
};

/** 描述电子表格范围，行列索引均为从 1 开始且包含结束位置。 */
export type SpreadsheetRange = {
  /** 范围起始行。 */
  startRow: number;
  /** 范围结束行。 */
  endRow: number;
  /** 范围起始列。 */
  startColumn: number;
  /** 范围结束列。 */
  endColumn: number;
};

/** 描述范围内显式记录的行尺寸。 */
export type SpreadsheetRowMetric = {
  /** 从 1 开始的全局行索引。 */
  index: number;
  /** 行高，单位为标准化渲染像素。 */
  height: number;
  /** 当前行是否隐藏。 */
  hidden: boolean;
};

/** 描述范围内显式记录的列尺寸。 */
export type SpreadsheetColumnMetric = {
  /** 从 1 开始的全局列索引。 */
  index: number;
  /** 列宽，单位为标准化渲染像素。 */
  width: number;
  /** 当前列是否隐藏。 */
  hidden: boolean;
};

/** 按需范围查询返回的稀疏工作表数据。 */
export type SpreadsheetRangeData = {
  /** 当前数据版本。 */
  revision: number;
  /** 实际返回的完整全局范围，可能因合并单元格而扩展。 */
  range: SpreadsheetRange;
  /** 范围内存在内容或样式的稀疏单元格。 */
  cells: readonly SpreadsheetCell[];
  /** 范围内的行尺寸。 */
  rows: readonly SpreadsheetRowMetric[];
  /** 范围内的列尺寸。 */
  columns: readonly SpreadsheetColumnMetric[];
  /** 与范围相交的完整合并区域。 */
  merges: readonly SpreadsheetMerge[];
  /** 与范围相交的图片。 */
  images: readonly SpreadsheetImage[];
  /** 与范围相交的图表。 */
  charts: readonly SpreadsheetChart[];
};
