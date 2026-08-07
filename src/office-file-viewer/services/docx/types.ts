import type { OfficeResourceSource } from '../resource-store/types';
import type { WordOutlineItem } from '../word/types';

/** DOCX 东亚标点字符间距压缩方式。 */
export type DocxCharacterSpacingControl =
  | 'doNotCompress'
  | 'compressPunctuation'
  | 'compressPunctuationAndJapaneseKana';

/** DOCX 标准化文档模型。 */
export type DocxDocument = {
  /** 面向用户展示的标题。 */
  title: string;
  /** 当前关联的页面模型。 */
  page: DocxPage;
  /** 按文档顺序排列的页面。 */
  pages?: DocxPageContent[];
  /** 按源文档顺序排列的内容块。 */
  blocks: DocxBlock[];
  /** 当前文档或页面包含的图片资源。 */
  images: DocxImage[];
  /** 源 DOCX 明确声明的大纲条目；为空时不显示目录侧栏。 */
  outline?: WordOutlineItem[];
  /** 是否保留源文档由节属性定义的物理分页。 */
  preserveSectionPagination?: boolean;
  /** 源文档声明的东亚标点字符间距压缩方式。 */
  characterSpacingControl?: DocxCharacterSpacingControl;
};

/** DOCX 单页内容及页眉页脚变体。 */
export type DocxPageContent = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 当前关联的页面模型。 */
  page: DocxPage;
  /** 是否保留源文档已定义的整页画布，避免测量分页再次拆开。 */
  preservePhysicalPage?: boolean;
  /** 按源文档顺序排列的内容块。 */
  blocks: DocxBlock[];
  /** 页眉在首页、偶数页和默认页上的内容变体。 */
  headers?: DocxPageRegionVariants<DocxBlock[]>;
  /** 页脚是否包含动态页码的首页、偶数页和默认页变体。 */
  footerPageNumbers?: DocxPageRegionVariants<boolean>;
  /** 是否启用首页独立页眉页脚。 */
  differentFirstPage?: boolean;
};

/** DOCX 首页、偶数页和默认页的页眉页脚变体。 */
export type DocxPageRegionVariants<T> = {
  /** 默认页使用的内容。 */
  default?: T;
  /** 首页使用的内容。 */
  first?: T;
  /** 偶数页使用的内容。 */
  even?: T;
};

/** DOCX 页面的布局与内容。 */
export type DocxPage = {
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 最小高度，单位为标准化渲染像素。 */
  minHeight: number;
  /** 上外边距，单位为标准化渲染像素。 */
  marginTop: number;
  /** 右外边距，单位为标准化渲染像素。 */
  marginRight: number;
  /** 下外边距，单位为标准化渲染像素。 */
  marginBottom: number;
  /** 左外边距，单位为标准化渲染像素。 */
  marginLeft: number;
  /** 页眉到页面顶部的距离，单位为标准化渲染像素。 */
  headerDistance?: number;
  /** 页脚到页面底部的距离，单位为标准化渲染像素。 */
  footerDistance?: number;
  /** 上边框的 CSS 样式。 */
  borderTop?: string;
  /** 右边框的 CSS 样式。 */
  borderRight?: string;
  /** 下边框的 CSS 样式。 */
  borderBottom?: string;
  /** 左边框的 CSS 样式。 */
  borderLeft?: string;
};

/** DOCX 对象的位置和定位基准。 */
export type DocxPosition = {
  /** 左侧位置或间距，单位由所属模型定义。 */
  left: number;
  /** 顶部位置或间距，单位由所属模型定义。 */
  top: number;
  /** 水平坐标采用的参考区域。 */
  relativeFromH?:
    | 'page'
    | 'margin'
    | 'column'
    | 'character'
    | 'leftMargin'
    | 'rightMargin'
    | 'insideMargin'
    | 'outsideMargin';
  /** 垂直坐标采用的参考区域。 */
  relativeFromV?:
    | 'page'
    | 'margin'
    | 'paragraph'
    | 'line'
    | 'text'
    | 'topMargin'
    | 'bottomMargin'
    | 'insideMargin'
    | 'outsideMargin';
  /** 层叠顺序，数值越大越靠近前景。 */
  zIndex?: number;
  /** 是否将对象放在正文内容之后。 */
  behindDoc?: boolean;
  /** 顺时针旋转角度，单位为度。 */
  rotation?: number;
  /** 是否沿水平方向翻转对象。 */
  flipH?: boolean;
  /** 是否沿垂直方向翻转对象。 */
  flipV?: boolean;
};

/** DOCX 内容块的内容与排版信息。 */
export type DocxBlock = DocxParagraphBlock | DocxTableBlock | DocxChartBlock;

/** DOCX 段落 内容块的内容与排版信息。 */
export type DocxParagraphBlock = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 分页拆分后仍指向原始正文块的稳定 ID。 */
  sourceBlockId?: string;
  /** 用于区分联合类型分支的类型标识。 */
  type: 'paragraph';
  /** 按源文档顺序排列的行内内容。 */
  inlines: DocxInline[];
  /** 文本内容。 */
  text: string;
  /** 源 DOCX 明确声明的大纲级别，使用从 0 开始的内部表示。 */
  outlineLevel?: number;
  /** 是否属于自动目录段落，用于还原目录制表位与打印样式。 */
  isTableOfContents?: boolean;
  /** 段落声明的制表位，位置单位为标准化渲染像素。 */
  tabStops?: DocxTabStop[];
  /** 水平对齐方式。 */
  align?: 'left' | 'center' | 'right' | 'justify';
  /** 行高，单位为标准化渲染像素。 */
  lineHeight?: number;
  /** 当前内容使用的渲染样式。 */
  style?: DocxTextStyle;
  /** 段后间距，单位为标准化渲染像素。 */
  spacingAfter?: number;
  /** 段前间距，单位为标准化渲染像素。 */
  spacingBefore?: number;
  /** 左缩进，单位为标准化渲染像素。 */
  indentLeft?: number;
  /** 右缩进，单位为标准化渲染像素。 */
  indentRight?: number;
  /** 首行缩进，单位为标准化渲染像素。 */
  firstLineIndent?: number;
  /** 背景颜色，使用 CSS 颜色值。 */
  backgroundColor?: string;
  /** 上边框的 CSS 样式。 */
  borderTop?: string;
  /** 右边框的 CSS 样式。 */
  borderRight?: string;
  /** 下边框的 CSS 样式。 */
  borderBottom?: string;
  /** 左边框的 CSS 样式。 */
  borderLeft?: string;
  /** 是否显式定义上边框。 */
  hasBorderTop?: boolean;
  /** 是否显式定义右边框。 */
  hasBorderRight?: boolean;
  /** 是否显式定义下边框。 */
  hasBorderBottom?: boolean;
  /** 是否显式定义左边框。 */
  hasBorderLeft?: boolean;
  /** 上内边距，单位为标准化渲染像素。 */
  paddingTop?: number;
  /** 右内边距，单位为标准化渲染像素。 */
  paddingRight?: number;
  /** 下内边距，单位为标准化渲染像素。 */
  paddingBottom?: number;
  /** 左内边距，单位为标准化渲染像素。 */
  paddingLeft?: number;
  /** 对象的定位信息及其参考坐标系。 */
  position?: DocxPosition;
};

/** DOCX 表格 内容块的内容与排版信息。 */
export type DocxTableBlock = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 分页拆分后仍指向原始正文块的稳定 ID。 */
  sourceBlockId?: string;
  /** 用于区分联合类型分支的类型标识。 */
  type: 'table';
  /** 按显示顺序排列的表格行。 */
  rows: DocxTableRow[];
  /** 宽度，单位为标准化渲染像素。 */
  width?: number;
  /** 水平对齐方式。 */
  align?: 'left' | 'center' | 'right';
  /** 按显示顺序排列的列定义。 */
  columns?: number[];
  /** 对象的定位信息及其参考坐标系。 */
  position?: DocxPosition;
  /** 上外边距，单位为标准化渲染像素。 */
  marginTop?: number;
  /** 是否位于形状内部；用于选择局部坐标和排版规则。 */
  insideShape?: boolean;
  /** 为匹配源文档视觉位置追加的顶部偏移，单位为标准化渲染像素。 */
  visualOffsetTop?: number;
};

/** DOCX 图表 内容块的内容与排版信息。 */
export type DocxChartBlock = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 分页拆分后仍指向原始正文块的稳定 ID。 */
  sourceBlockId?: string;
  /** 用于区分联合类型分支的类型标识。 */
  type: 'chart';
  /** 图表渲染相关文案。 */
  chart: import('../../shared/ooxml/charts').OfficeChartModel;
  /** WPS 静态图表快照可在页面挂载后按需读取。 */
  snapshotSource?: OfficeResourceSource;
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 对象的定位信息及其参考坐标系。 */
  position?: DocxPosition;
};

/** DOCX 表格行及其高度规则。 */
export type DocxTableRow = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 按显示顺序排列的单元格。 */
  cells: DocxTableCell[];
  /** 高度，单位为标准化渲染像素。 */
  height?: number;
  /** 表格行高采用自动、最小值或固定值的规则。 */
  heightRule?: 'auto' | 'atLeast' | 'exact';
};

/** DOCX 表格单元格的内容与样式。 */
export type DocxTableCell = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 按源文档顺序排列的内容块。 */
  blocks: DocxBlock[];
  /** 表格单元格横向跨越的列数。 */
  colSpan?: number;
  /** 表格单元格纵向跨越的行数。 */
  rowSpan?: number;
  /** 宽度，单位为标准化渲染像素。 */
  width?: number;
  /** 垂直对齐方式。 */
  verticalAlign?: 'top' | 'middle' | 'bottom';
  /** 背景颜色，使用 CSS 颜色值。 */
  backgroundColor?: string;
  /** 上边框的 CSS 样式。 */
  borderTop?: string;
  /** 右边框的 CSS 样式。 */
  borderRight?: string;
  /** 下边框的 CSS 样式。 */
  borderBottom?: string;
  /** 左边框的 CSS 样式。 */
  borderLeft?: string;
  /** 是否显式定义上边框。 */
  hasBorderTop?: boolean;
  /** 是否显式定义右边框。 */
  hasBorderRight?: boolean;
  /** 是否显式定义下边框。 */
  hasBorderBottom?: boolean;
  /** 是否显式定义左边框。 */
  hasBorderLeft?: boolean;
  /** 上内边距，单位为标准化渲染像素。 */
  paddingTop?: number;
  /** 右内边距，单位为标准化渲染像素。 */
  paddingRight?: number;
  /** 下内边距，单位为标准化渲染像素。 */
  paddingBottom?: number;
  /** 左内边距，单位为标准化渲染像素。 */
  paddingLeft?: number;
  /** 是否禁止单元格内容自动换行。 */
  noWrap?: boolean;
};

/** DOCX 行内内容模型。 */
export type DocxInline =
  | DocxTextInline
  | DocxTabInline
  | DocxBreakInline
  | DocxImageInline
  | DocxChartInline
  | DocxShapeInline;

/** DOCX 文本 行内内容模型。 */
export type DocxTextInline = {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'text';
  /** 文本内容。 */
  text: string;
  /** 当前内容使用的渲染样式。 */
  style?: DocxTextStyle;
};

/** DOCX 制表符 行内内容模型。 */
export type DocxTabInline = {
  /** 用于区分制表符与普通文本。 */
  type: 'tab';
  /** 制表符沿用的文字样式。 */
  style?: DocxTextStyle;
};

/** DOCX 制表符 停止点。 */
export type DocxTabStop = {
  /** 制表位相对段落起点的位置，单位为标准化渲染像素。 */
  position: number;
  /** 制表位的对齐方式。 */
  align: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'number';
  /** 制表位引导符。 */
  leader?: 'dot' | 'hyphen' | 'underscore' | 'middleDot' | 'none';
};

/** DOCX 换行 行内内容模型。 */
export type DocxBreakInline = {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'break';
};

/** DOCX 图片 行内内容模型。 */
export type DocxImageInline = {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'image';
  /** 当前关联的图片资源或图片模型。 */
  image: DocxImage;
};

/** DOCX 图表 行内内容模型。 */
export type DocxChartInline = {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'chart';
  /** 图表渲染相关文案。 */
  chart: DocxChartBlock;
};

/** DOCX 形状 行内内容模型。 */
export type DocxShapeInline = {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'shape';
  /** 当前关联的形状模型。 */
  shape: DocxShape;
};

/** DOCX 形状的尺寸、位置和子项。 */
export type DocxShape = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 对象的定位信息及其参考坐标系。 */
  position?: DocxPosition;
  /** 按显示顺序排列的项目。 */
  items: DocxShapeItem[];
};

/** DOCX 形状子项的几何、样式和内容。 */
export type DocxShapeItem = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 当前模型对应的 Office 内容类型。 */
  kind: 'rect' | 'ellipse' | 'line' | 'path';
  /** 左侧位置或间距，单位由所属模型定义。 */
  left: number;
  /** 顶部位置或间距，单位由所属模型定义。 */
  top: number;
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 上内边距，单位为标准化渲染像素。 */
  paddingTop?: number;
  /** 右内边距，单位为标准化渲染像素。 */
  paddingRight?: number;
  /** 下内边距，单位为标准化渲染像素。 */
  paddingBottom?: number;
  /** 左内边距，单位为标准化渲染像素。 */
  paddingLeft?: number;
  /** 在压缩包、复合文档或资源表中的路径。 */
  path?: string;
  /** 矢量路径使用的坐标范围。 */
  viewBox?: string;
  /** 填充颜色，使用 CSS 颜色值。 */
  fillColor?: string;
  /** 形状背景图可直接使用 URL，也可在页面挂载后按需读取。 */
  imageSrc?: OfficeResourceSource;
  /** 边框。 */
  border?: string;
  /** 轮廓颜色，使用 CSS 颜色值。 */
  strokeColor?: string;
  /** 轮廓宽度，单位为标准化渲染像素。 */
  strokeWidth?: number;
  /** 轮廓虚线各段长度的 CSS 配置。 */
  strokeDasharray?: string;
  /** 圆角半径；数值单位为标准化渲染像素。 */
  borderRadius?: number | string;
  /** 形状内部文字的垂直对齐方式。 */
  textVerticalAlign?: 'top' | 'middle' | 'bottom';
  /** 是否根据文本内容自动调整形状尺寸。 */
  fitShapeToText?: boolean;
  /** 是否禁止单元格内容自动换行。 */
  noWrap?: boolean;
  /** 按源文档顺序排列的内容块。 */
  blocks?: DocxBlock[];
  /** 按源文档顺序排列的段落。 */
  paragraphs?: DocxParagraphBlock[];
};

/** DOCX 图片资源及其显示信息。 */
export type DocxImage = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 面向用户展示的名称。 */
  name?: string;
  /** 图片无法显示时使用的替代文本。 */
  alt?: string;
  /** 图片可直接使用 URL，也可在页面挂载后按需读取。 */
  src: OfficeResourceSource;
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 对象的定位信息及其参考坐标系。 */
  position?: DocxPosition;
};

/** DOCX 文本渲染样式。 */
export type DocxTextStyle = {
  /** 是否使用粗体。 */
  bold?: boolean;
  /** 是否使用斜体。 */
  italic?: boolean;
  /** 是否绘制下划线。 */
  underline?: boolean;
  /** 使用的删除线类型。 */
  strike?: boolean;
  /** 是否使用小型大写字母样式。 */
  smallCaps?: boolean;
  /** 是否将文字显示为全大写。 */
  allCaps?: boolean;
  /** 前景或文字颜色，使用 CSS 颜色值。 */
  color?: string;
  /** 字号，单位为标准化渲染像素。 */
  fontSize?: number;
  /** 字体族名称。 */
  fontFamily?: string;
  /** 水平对齐方式。 */
  align?: 'left' | 'center' | 'right' | 'justify';
  /** 行高，单位为标准化渲染像素。 */
  lineHeight?: number;
  /** 段前间距，单位为标准化渲染像素。 */
  spacingBefore?: number;
  /** 段后间距，单位为标准化渲染像素。 */
  spacingAfter?: number;
  /** 左缩进，单位为标准化渲染像素。 */
  indentLeft?: number;
  /** 右缩进，单位为标准化渲染像素。 */
  indentRight?: number;
  /** 首行缩进，单位为标准化渲染像素。 */
  firstLineIndent?: number;
  /** 背景颜色，使用 CSS 颜色值。 */
  backgroundColor?: string;
  /** 上边框的 CSS 样式。 */
  borderTop?: string;
  /** 右边框的 CSS 样式。 */
  borderRight?: string;
  /** 下边框的 CSS 样式。 */
  borderBottom?: string;
  /** 左边框的 CSS 样式。 */
  borderLeft?: string;
  /** 上内边距，单位为标准化渲染像素。 */
  paddingTop?: number;
  /** 右内边距，单位为标准化渲染像素。 */
  paddingRight?: number;
  /** 下内边距，单位为标准化渲染像素。 */
  paddingBottom?: number;
  /** 左内边距，单位为标准化渲染像素。 */
  paddingLeft?: number;
};
