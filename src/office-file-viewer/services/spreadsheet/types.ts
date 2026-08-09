import type { OfficeChartModel } from '../../shared/ooxml/charts';
import type { OfficeHyperlink } from '../../shared/hyperlink';
import type { OfficeResourceSource } from '../resource-store';

/** 描述电子表格标准模型过程中可继续处理的警告。 */
export type SpreadsheetWarning = {
  /** 供程序识别当前情况的稳定代码。 */
  code: string;
  /** 面向调用方或用户展示的说明。 */
  message: string;
  /** 产生警告的工作表名称。 */
  sheetName?: string;
  /** 在所属数据范围中的偏移位置。 */
  offset?: number;
};

/** 包含工作表和资源的标准化工作簿。 */
export type SpreadsheetWorkbook = {
  /** 按工作簿顺序排列的工作表。 */
  sheets: SpreadsheetSheet[];
  /** 工作簿级定义名称到静态目标地址的映射。 */
  definedNames?: Record<string, string>;
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

/** 单张工作表的网格、合并区域及浮动对象。 */
export type SpreadsheetSheet = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 面向用户展示的名称。 */
  name: string;
  /** 在压缩包、复合文档或资源表中的路径。 */
  path: string;
  /** 当前模型对应的 Office 内容类型。 */
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
  /** 按列号排列的工作表列。 */
  columns: SpreadsheetColumn[];
  /** 按行号排列的工作表行。 */
  rows: SpreadsheetRow[];
  /** 工作表声明的合并单元格区域。 */
  merges: SpreadsheetMerge[];
  /** 工作表包含的浮动图片。 */
  images: SpreadsheetImage[];
  /** 工作表包含的浮动图表。 */
  charts: SpreadsheetChart[];
  /** 工作表声明的稀疏超链接范围。 */
  hyperlinks?: SpreadsheetHyperlinkRange[];
};

/** 工作表中一个可绑定到单元格窗口的超链接范围。 */
export type SpreadsheetHyperlinkRange = {
  /** 源文件声明的 A1 范围。 */
  ref: string;
  /** 范围起始行索引。 */
  startRow: number;
  /** 范围起始列索引。 */
  startColumn: number;
  /** 范围结束行索引。 */
  endRow: number;
  /** 范围结束列索引。 */
  endColumn: number;
  /** 范围内单元格共享的标准超链接。 */
  hyperlink: OfficeHyperlink;
};

/** 工作表列的位置、宽度和隐藏状态。 */
export type SpreadsheetColumn = {
  /** 在所属集合中的零基索引。 */
  index: number;
  /** 面向用户展示的标签文本。 */
  label: string;
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 是否隐藏当前项目。 */
  hidden?: boolean;
};

/** 工作表行的位置、高度、隐藏状态和单元格。 */
export type SpreadsheetRow = {
  /** 在所属集合中的零基索引。 */
  index: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 是否由源文件显式固定行高；否则允许换行内容按 Excel 规则自动撑高。 */
  customHeight?: boolean;
  /** 是否隐藏当前项目。 */
  hidden?: boolean;
  /** 按显示顺序排列的单元格。 */
  cells: SpreadsheetCell[];
};

/** 工作表单元格的值、公式、类型和样式。 */
export type SpreadsheetCell = {
  /** 单元格 A1 引用，例如 C5。 */
  ref: string;
  /** 单元格或记录所在的行索引。 */
  rowIndex: number;
  /** 单元格或记录所在的列索引。 */
  columnIndex: number;
  /** 单元格格式化后的显示文本。 */
  value: string;
  /** 格式化前从源文件读取的原始值。 */
  rawValue?: string;
  /** 源文件中的单元格值类型标识。 */
  type?: string;
  /** 单元格引用的样式表索引。 */
  styleId?: number;
  /** 当前内容使用的渲染样式。 */
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
  /** 源文件为单元格声明的超链接。 */
  hyperlink?: OfficeHyperlink;
};

/** 工作表合并区域及其行列边界。 */
export type SpreadsheetMerge = {
  /** 合并区域的 A1 引用文本。 */
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

/** 浮动对象锚点对应的单元格和偏移比例。 */
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

/** 工作表浮动图片的资源、锚点和显示尺寸。 */
export type SpreadsheetImage = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 面向用户展示的名称。 */
  name?: string;
  /** 图片可直接用于 img 元素的资源地址。 */
  src: string | OfficeResourceSource;
  /** 图片无法显示时使用的替代文本。 */
  alt?: string;
  /** 起始锚点。 */
  from: SpreadsheetAnchorPoint;
  /** 结束锚点。 */
  to: SpreadsheetAnchorPoint;
  /** 相对定位区域左侧的横坐标，单位为标准化渲染像素。 */
  x: number;
  /** 相对定位区域顶部的纵坐标，单位为标准化渲染像素。 */
  y: number;
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 源文件为图片声明的超链接。 */
  hyperlink?: OfficeHyperlink;
};

/** 工作表浮动图表的模型、锚点和显示尺寸。 */
export type SpreadsheetChart = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 面向用户展示的标题。 */
  title?: string;
  /** 当前浮动对象承载的标准图表模型。 */
  chart: OfficeChartModel;
  /** 起始锚点。 */
  from: SpreadsheetAnchorPoint;
  /** 结束锚点。 */
  to: SpreadsheetAnchorPoint;
  /** 相对定位区域左侧的横坐标，单位为标准化渲染像素。 */
  x: number;
  /** 相对定位区域顶部的纵坐标，单位为标准化渲染像素。 */
  y: number;
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 源文件为图表对象声明的超链接。 */
  hyperlink?: OfficeHyperlink;
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
  /** 是否使用粗体。 */
  bold?: boolean;
  /** 是否使用斜体。 */
  italic?: boolean;
  /** 是否绘制下划线。 */
  underline?: boolean;
  /** 前景或文字颜色，使用 CSS 颜色值。 */
  color?: string;
  /** 字体族名称。 */
  fontFamily?: string;
  /** 字号，单位为标准化渲染像素。 */
  fontSize?: number;
  /** 背景颜色，使用 CSS 颜色值。 */
  backgroundColor?: string;
  /** 单元格内容的水平对齐方式。 */
  horizontalAlign?: 'left' | 'center' | 'right' | 'justify';
  /** 垂直对齐方式。 */
  verticalAlign?: 'top' | 'middle' | 'bottom';
  /** 是否允许单元格文本自动换行。 */
  wrapText?: boolean;
  /** 是否在单元格宽度不足时缩小文字以适应边界。 */
  shrinkToFit?: boolean;
  /** 是否存在任意方向的有效边框。 */
  border?: boolean;
  /** 上边框的 CSS 样式。 */
  borderTop?: string;
  /** 右边框的 CSS 样式。 */
  borderRight?: string;
  /** 下边框的 CSS 样式。 */
  borderBottom?: string;
  /** 左边框的 CSS 样式。 */
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
  /** 是否由源文件显式固定行高；否则为自动行高。 */
  customHeight?: boolean;
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
