import type { WordOutlineItem } from '../word/types';

/** 描述 DOCX 解析生成的标准化文档模型。 */
export type DocxDocument = {
  /** DocxDocument 对外展示的标题。 */
  title: string;
  /** DocxDocument 当前关联的页面模型。 */
  page: DocxPage;
  /** DocxDocument 包含的 pages 有序集合。 */
  pages?: DocxPageContent[];
  /** DocxDocument 包含的 blocks 有序集合。 */
  blocks: DocxBlock[];
  /** DocxDocument 包含的 images 有序集合。 */
  images: DocxImage[];
  /** 源 DOCX 明确声明的大纲条目；为空时不显示目录侧栏。 */
  outline?: WordOutlineItem[];
  /** 是否保留源文档由节属性定义的物理分页。 */
  preserveSectionPagination?: boolean;
};

/** 描述 DocxPageContent 在 DOCX 解析中的数据结构。 */
export type DocxPageContent = {
  /** DocxPageContent 在所属文档或任务中的唯一标识。 */
  id: string;
  /** DocxPageContent 当前关联的页面模型。 */
  page: DocxPage;
  /** DocxPageContent 包含的 blocks 有序集合。 */
  blocks: DocxBlock[];
  /** 页眉在首页、偶数页和默认页上的内容变体。 */
  headers?: DocxPageRegionVariants<DocxBlock[]>;
  /** 页脚是否包含动态页码的首页、偶数页和默认页变体。 */
  footerPageNumbers?: DocxPageRegionVariants<boolean>;
  /** 是否启用首页独立页眉页脚。 */
  differentFirstPage?: boolean;
};

/** 描述 DOCX 页眉页脚针对不同页型的内容变体。 */
export type DocxPageRegionVariants<T> = {
  /** 默认页使用的内容。 */
  default?: T;
  /** 首页使用的内容。 */
  first?: T;
  /** 偶数页使用的内容。 */
  even?: T;
};

/** 描述 DocxPage 在 DOCX 解析中的数据结构。 */
export type DocxPage = {
  /** DocxPage 的 width 尺寸或坐标，单位为标准化渲染像素。 */
  width: number;
  /** DocxPage 的 minHeight 尺寸或坐标，单位为标准化渲染像素。 */
  minHeight: number;
  /** DocxPage 的对应间距，单位为标准化渲染像素。 */
  marginTop: number;
  /** DocxPage 的对应间距，单位为标准化渲染像素。 */
  marginRight: number;
  /** DocxPage 的对应间距，单位为标准化渲染像素。 */
  marginBottom: number;
  /** DocxPage 的对应间距，单位为标准化渲染像素。 */
  marginLeft: number;
  /** 页眉到页面顶部的距离，单位为标准化渲染像素。 */
  headerDistance?: number;
  /** 页脚到页面底部的距离，单位为标准化渲染像素。 */
  footerDistance?: number;
  /** DocxPage 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderTop?: string;
  /** DocxPage 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderRight?: string;
  /** DocxPage 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderBottom?: string;
  /** DocxPage 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderLeft?: string;
};

/** 描述 DOCX 解析对象的位置和定位基准。 */
export type DocxPosition = {
  /** DocxPosition 的 left 尺寸或坐标，单位为标准化渲染像素。 */
  left: number;
  /** DocxPosition 的 top 尺寸或坐标，单位为标准化渲染像素。 */
  top: number;
  /** 水平坐标采用的参考区域；未提供时沿用来源格式或渲染器的默认规则。 */
  relativeFromH?:
    | 'page'
    | 'margin'
    | 'column'
    | 'character'
    | 'leftMargin'
    | 'rightMargin'
    | 'insideMargin'
    | 'outsideMargin';
  /** 垂直坐标采用的参考区域；未提供时沿用来源格式或渲染器的默认规则。 */
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
  /** DocxPosition 的层叠顺序，数值越大越靠近前景。 */
  zIndex?: number;
  /** 是否将对象放在正文内容之后。 */
  behindDoc?: boolean;
  /** DocxPosition 的顺时针旋转角度，单位为度；未提供时沿用来源格式或渲染器的默认规则。 */
  rotation?: number;
  /** 是否沿水平方向翻转对象。 */
  flipH?: boolean;
  /** 是否沿垂直方向翻转对象。 */
  flipV?: boolean;
};

/** 描述 DocxBlock 在 DOCX 解析中的数据结构。 */
export type DocxBlock = DocxParagraphBlock | DocxTableBlock | DocxChartBlock;

/** 描述 DocxParagraphBlock 在 DOCX 解析中的数据结构。 */
export type DocxParagraphBlock = {
  /** DocxParagraphBlock 在所属文档或任务中的唯一标识。 */
  id: string;
  /** 用于区分 DocxParagraphBlock 不同结构分支的类型标识。 */
  type: 'paragraph';
  /** DocxParagraphBlock 包含的 inlines 有序集合。 */
  inlines: DocxInline[];
  /** DocxParagraphBlock 携带或渲染的文本内容。 */
  text: string;
  /** 源 DOCX 明确声明的大纲级别，使用从 0 开始的内部表示。 */
  outlineLevel?: number;
  /** DocxParagraphBlock 的水平对齐方式；未提供时沿用来源格式或渲染器的默认规则。 */
  align?: 'left' | 'center' | 'right' | 'justify';
  /** DocxParagraphBlock 的行高，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  lineHeight?: number;
  /** DocxParagraphBlock 使用的渲染或文本样式。 */
  style?: DocxTextStyle;
  /** DocxParagraphBlock 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  spacingAfter?: number;
  /** DocxParagraphBlock 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  spacingBefore?: number;
  /** DocxParagraphBlock 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  indentLeft?: number;
  /** DocxParagraphBlock 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  indentRight?: number;
  /** DocxParagraphBlock 的首行缩进，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  firstLineIndent?: number;
  /** DocxParagraphBlock 的背景颜色，使用 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
  backgroundColor?: string;
  /** DocxParagraphBlock 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderTop?: string;
  /** DocxParagraphBlock 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderRight?: string;
  /** DocxParagraphBlock 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderBottom?: string;
  /** DocxParagraphBlock 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderLeft?: string;
  /** 是否显式定义 DocxParagraphBlock 对应方向的边框。 */
  hasBorderTop?: boolean;
  /** 是否显式定义 DocxParagraphBlock 对应方向的边框。 */
  hasBorderRight?: boolean;
  /** 是否显式定义 DocxParagraphBlock 对应方向的边框。 */
  hasBorderBottom?: boolean;
  /** 是否显式定义 DocxParagraphBlock 对应方向的边框。 */
  hasBorderLeft?: boolean;
  /** DocxParagraphBlock 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingTop?: number;
  /** DocxParagraphBlock 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingRight?: number;
  /** DocxParagraphBlock 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingBottom?: number;
  /** DocxParagraphBlock 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingLeft?: number;
  /** DocxParagraphBlock 的定位信息及其参考坐标系。 */
  position?: DocxPosition;
};

/** 描述 DocxTableBlock 在 DOCX 解析中的数据结构。 */
export type DocxTableBlock = {
  /** DocxTableBlock 在所属文档或任务中的唯一标识。 */
  id: string;
  /** 用于区分 DocxTableBlock 不同结构分支的类型标识。 */
  type: 'table';
  /** DocxTableBlock 包含的 rows 有序集合。 */
  rows: DocxTableRow[];
  /** DocxTableBlock 的 width 尺寸或坐标，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  width?: number;
  /** DocxTableBlock 的水平对齐方式；未提供时沿用来源格式或渲染器的默认规则。 */
  align?: 'left' | 'center' | 'right';
  /** DocxTableBlock 包含的 columns 有序集合。 */
  columns?: number[];
  /** DocxTableBlock 的定位信息及其参考坐标系。 */
  position?: DocxPosition;
  /** DocxTableBlock 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  marginTop?: number;
  /** 是否位于形状内部；用于选择局部坐标和排版规则。 */
  insideShape?: boolean;
  /** DocxTableBlock 为匹配源文档视觉位置追加的顶部偏移，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  visualOffsetTop?: number;
};

/** 描述 DocxChartBlock 在 DOCX 解析中的数据结构。 */
export type DocxChartBlock = {
  /** DocxChartBlock 在所属文档或任务中的唯一标识。 */
  id: string;
  /** 用于区分 DocxChartBlock 不同结构分支的类型标识。 */
  type: 'chart';
  /** DocxChartBlock 当前关联的图表模型。 */
  chart: import('../../shared/ooxml/charts').OfficeChartModel;
  /** DocxChartBlock 的 width 尺寸或坐标，单位为标准化渲染像素。 */
  width: number;
  /** DocxChartBlock 的 height 尺寸或坐标，单位为标准化渲染像素。 */
  height: number;
  /** DocxChartBlock 的定位信息及其参考坐标系。 */
  position?: DocxPosition;
};

/** 描述 DocxTableRow 在 DOCX 解析中的数据结构。 */
export type DocxTableRow = {
  /** DocxTableRow 在所属文档或任务中的唯一标识。 */
  id: string;
  /** DocxTableRow 包含的 cells 有序集合。 */
  cells: DocxTableCell[];
  /** DocxTableRow 的 height 尺寸或坐标，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  height?: number;
  /** DocxTableRow 关联的 heightRule 结构；字段形状由 'auto' | 'atLeast' | 'exact' 定义；未提供时使用来源格式或渲染器的默认行为。 */
  heightRule?: 'auto' | 'atLeast' | 'exact';
};

/** 描述 DocxTableCell 在 DOCX 解析中的数据结构。 */
export type DocxTableCell = {
  /** DocxTableCell 在所属文档或任务中的唯一标识。 */
  id: string;
  /** DocxTableCell 包含的 blocks 有序集合。 */
  blocks: DocxBlock[];
  /** 表格单元格横向跨越的列数。 */
  colSpan?: number;
  /** 表格单元格纵向跨越的行数。 */
  rowSpan?: number;
  /** DocxTableCell 的 width 尺寸或坐标，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  width?: number;
  /** DocxTableCell 的垂直对齐方式；未提供时沿用来源格式或渲染器的默认规则。 */
  verticalAlign?: 'top' | 'middle' | 'bottom';
  /** DocxTableCell 的背景颜色，使用 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
  backgroundColor?: string;
  /** DocxTableCell 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderTop?: string;
  /** DocxTableCell 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderRight?: string;
  /** DocxTableCell 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderBottom?: string;
  /** DocxTableCell 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderLeft?: string;
  /** 是否显式定义 DocxTableCell 对应方向的边框。 */
  hasBorderTop?: boolean;
  /** 是否显式定义 DocxTableCell 对应方向的边框。 */
  hasBorderRight?: boolean;
  /** 是否显式定义 DocxTableCell 对应方向的边框。 */
  hasBorderBottom?: boolean;
  /** 是否显式定义 DocxTableCell 对应方向的边框。 */
  hasBorderLeft?: boolean;
  /** DocxTableCell 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingTop?: number;
  /** DocxTableCell 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingRight?: number;
  /** DocxTableCell 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingBottom?: number;
  /** DocxTableCell 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingLeft?: number;
  /** 是否禁止单元格内容自动换行。 */
  noWrap?: boolean;
};

/** 描述 DocxInline 在 DOCX 解析中的数据结构。 */
export type DocxInline =
  | DocxTextInline
  | DocxBreakInline
  | DocxImageInline
  | DocxChartInline
  | DocxShapeInline;

/** 描述 DocxTextInline 在 DOCX 解析中的数据结构。 */
export type DocxTextInline = {
  /** 用于区分 DocxTextInline 不同结构分支的类型标识。 */
  type: 'text';
  /** DocxTextInline 携带或渲染的文本内容。 */
  text: string;
  /** DocxTextInline 使用的渲染或文本样式。 */
  style?: DocxTextStyle;
};

/** 描述 DocxBreakInline 在 DOCX 解析中的数据结构。 */
export type DocxBreakInline = {
  /** 用于区分 DocxBreakInline 不同结构分支的类型标识。 */
  type: 'break';
};

/** 描述 DocxImageInline 在 DOCX 解析中的数据结构。 */
export type DocxImageInline = {
  /** 用于区分 DocxImageInline 不同结构分支的类型标识。 */
  type: 'image';
  /** DocxImageInline 当前关联的图片资源或图片模型。 */
  image: DocxImage;
};

/** 描述 DocxChartInline 在 DOCX 解析中的数据结构。 */
export type DocxChartInline = {
  /** 用于区分 DocxChartInline 不同结构分支的类型标识。 */
  type: 'chart';
  /** DocxChartInline 当前关联的图表模型。 */
  chart: DocxChartBlock;
};

/** 描述 DocxShapeInline 在 DOCX 解析中的数据结构。 */
export type DocxShapeInline = {
  /** 用于区分 DocxShapeInline 不同结构分支的类型标识。 */
  type: 'shape';
  /** DocxShapeInline 关联的 shape 结构；字段形状由 DocxShape 定义。 */
  shape: DocxShape;
};

/** 描述 DocxShape 在 DOCX 解析中的数据结构。 */
export type DocxShape = {
  /** DocxShape 在所属文档或任务中的唯一标识。 */
  id: string;
  /** DocxShape 的 width 尺寸或坐标，单位为标准化渲染像素。 */
  width: number;
  /** DocxShape 的 height 尺寸或坐标，单位为标准化渲染像素。 */
  height: number;
  /** DocxShape 的定位信息及其参考坐标系。 */
  position?: DocxPosition;
  /** DocxShape 包含的 items 有序集合。 */
  items: DocxShapeItem[];
};

/** 描述 DocxShapeItem 在 DOCX 解析中的数据结构。 */
export type DocxShapeItem = {
  /** DocxShapeItem 在所属文档或任务中的唯一标识。 */
  id: string;
  /** 标识 DocxShapeItem 对应的 Office 文件或数据种类。 */
  kind: 'rect' | 'ellipse' | 'line' | 'path';
  /** DocxShapeItem 的 left 尺寸或坐标，单位为标准化渲染像素。 */
  left: number;
  /** DocxShapeItem 的 top 尺寸或坐标，单位为标准化渲染像素。 */
  top: number;
  /** DocxShapeItem 的 width 尺寸或坐标，单位为标准化渲染像素。 */
  width: number;
  /** DocxShapeItem 的 height 尺寸或坐标，单位为标准化渲染像素。 */
  height: number;
  /** DocxShapeItem 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingTop?: number;
  /** DocxShapeItem 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingRight?: number;
  /** DocxShapeItem 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingBottom?: number;
  /** DocxShapeItem 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingLeft?: number;
  /** DocxShapeItem 在压缩包、复合文档或图形数据中的路径。 */
  path?: string;
  /** DocxShapeItem 的 viewBox 文本值。 */
  viewBox?: string;
  /** DocxShapeItem 的 fillColor 文本值。 */
  fillColor?: string;
  /** DocxShapeItem 的 imageSrc 文本值。 */
  imageSrc?: string;
  /** DocxShapeItem 的 border 文本值。 */
  border?: string;
  /** DocxShapeItem 的 strokeColor 文本值。 */
  strokeColor?: string;
  /** DocxShapeItem 的轮廓宽度，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  strokeWidth?: number;
  /** DocxShapeItem 的 strokeDasharray 文本值。 */
  strokeDasharray?: string;
  /** DocxShapeItem 关联的 borderRadius 结构；字段形状由 number | string 定义；未提供时使用来源格式或渲染器的默认行为。 */
  borderRadius?: number | string;
  /** DocxShapeItem 关联的 textVerticalAlign 结构；字段形状由 'top' | 'middle' | 'bottom' 定义；未提供时使用来源格式或渲染器的默认行为。 */
  textVerticalAlign?: 'top' | 'middle' | 'bottom';
  /** 是否启用 fitShapeToText 对应的格式选项；未提供时使用来源格式或渲染器的默认行为。 */
  fitShapeToText?: boolean;
  /** 是否禁止单元格内容自动换行。 */
  noWrap?: boolean;
  /** DocxShapeItem 包含的 blocks 有序集合。 */
  blocks?: DocxBlock[];
  /** DocxShapeItem 包含的 paragraphs 有序集合。 */
  paragraphs?: DocxParagraphBlock[];
};

/** 描述 DocxImage 在 DOCX 解析中的数据结构。 */
export type DocxImage = {
  /** DocxImage 在所属文档或任务中的唯一标识。 */
  id: string;
  /** DocxImage 的可读名称。 */
  name?: string;
  /** DocxImage 的 alt 文本值。 */
  alt?: string;
  /** DocxImage 的 src 文本值。 */
  src: string;
  /** DocxImage 的 width 尺寸或坐标，单位为标准化渲染像素。 */
  width: number;
  /** DocxImage 的 height 尺寸或坐标，单位为标准化渲染像素。 */
  height: number;
  /** DocxImage 的定位信息及其参考坐标系。 */
  position?: DocxPosition;
};

/** 描述 DOCX 解析使用的样式参数。 */
export type DocxTextStyle = {
  /** 是否使用粗体渲染 DocxTextStyle；未提供时沿用来源格式或渲染器的默认规则。 */
  bold?: boolean;
  /** 是否使用斜体渲染 DocxTextStyle；未提供时沿用来源格式或渲染器的默认规则。 */
  italic?: boolean;
  /** 是否为 DocxTextStyle 绘制下划线；未提供时沿用来源格式或渲染器的默认规则。 */
  underline?: boolean;
  /** DocxTextStyle 使用的删除线类型；未提供时沿用来源格式或渲染器的默认规则。 */
  strike?: boolean;
  /** 是否启用 smallCaps 对应的格式选项；未提供时使用来源格式或渲染器的默认行为。 */
  smallCaps?: boolean;
  /** 是否启用 allCaps 对应的格式选项；未提供时使用来源格式或渲染器的默认行为。 */
  allCaps?: boolean;
  /** DocxTextStyle 的前景或文本颜色，使用标准化 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
  color?: string;
  /** DocxTextStyle 的字号，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  fontSize?: number;
  /** DocxTextStyle 的字体族名称；未提供时沿用来源格式或渲染器的默认规则。 */
  fontFamily?: string;
  /** DocxTextStyle 的水平对齐方式；未提供时沿用来源格式或渲染器的默认规则。 */
  align?: 'left' | 'center' | 'right' | 'justify';
  /** DocxTextStyle 的行高，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  lineHeight?: number;
  /** DocxTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  spacingBefore?: number;
  /** DocxTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  spacingAfter?: number;
  /** DocxTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  indentLeft?: number;
  /** DocxTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  indentRight?: number;
  /** DocxTextStyle 的首行缩进，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  firstLineIndent?: number;
  /** DocxTextStyle 的背景颜色，使用 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
  backgroundColor?: string;
  /** DocxTextStyle 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderTop?: string;
  /** DocxTextStyle 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderRight?: string;
  /** DocxTextStyle 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderBottom?: string;
  /** DocxTextStyle 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderLeft?: string;
  /** DocxTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingTop?: number;
  /** DocxTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingRight?: number;
  /** DocxTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingBottom?: number;
  /** DocxTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingLeft?: number;
};
