/** 描述 DOC 二进制解析生成的标准化文档模型。 */
export type DocDocument = {
  /** DocDocument 对外展示的标题。 */
  title: string;
  /** DocDocument 当前关联的页面模型。 */
  page: DocPage;
  /** DocDocument 包含的 blocks 有序集合。 */
  blocks: DocBlock[];
  /** DocDocument 包含的 paragraphs 有序集合。 */
  paragraphs: DocParagraph[];
  /** DocDocument 包含的 images 有序集合。 */
  images: DocImage[];
  /** 从页眉 story 中识别出的徽标图片；未提供时不渲染页眉图片。 */
  headerImage?: DocImage;
  /** 页脚 story 是否包含 PAGE 字段。 */
  footerPageNumbers?: boolean;
  /** DocDocument 解析时产生但不阻止继续预览的警告集合。 */
  warnings: string[];
  /** DOC/WPS 文档持有且需要在销毁时释放的浏览器资源。 */
  resources?: DocResources;
};

/** 记录 DOC/WPS 文档持有且需要统一管理的资源。 */
export type DocResources = {
  /** 浏览器创建的对象 URL 集合，文档释放时必须逐一撤销。 */
  objectUrls: string[];
};

/** 释放 DOC/WPS 文档创建的 Blob URL；重复调用保持幂等。 */
export function disposeDocDocument(document: DocDocument | undefined) {
  const urls = document?.resources?.objectUrls;
  if (!urls?.length) return;
  const uniqueUrls = new Set(urls);
  urls.length = 0;
  if (typeof URL === 'undefined' || typeof URL.revokeObjectURL !== 'function') {
    return;
  }
  uniqueUrls.forEach((url) => URL.revokeObjectURL(url));
}

/** 描述 DocPage 在 DOC 二进制解析中的数据结构。 */
export type DocPage = {
  /** 标准化页面宽度，单位为渲染像素。 */
  width: number;
  /** 标准化页面最小高度，单位为渲染像素。 */
  minHeight: number;
  /** DocPage 的对应间距，单位为标准化渲染像素。 */
  marginTop: number;
  /** DocPage 的对应间距，单位为标准化渲染像素。 */
  marginRight: number;
  /** DocPage 的对应间距，单位为标准化渲染像素。 */
  marginBottom: number;
  /** DocPage 的对应间距，单位为标准化渲染像素。 */
  marginLeft: number;
};

/** 描述 DocParagraph 在 DOC 二进制解析中的数据结构。 */
export type DocParagraph = {
  /** DocParagraph 在所属文档或任务中的唯一标识。 */
  id: string;
  /** DocParagraph 携带或渲染的文本内容。 */
  text: string;
};

/** 描述 DocBlock 在 DOC 二进制解析中的数据结构。 */
export type DocBlock = DocParagraphBlock | DocTableBlock | DocListBlock;

/** 描述 DocParagraphBlock 在 DOC 二进制解析中的数据结构。 */
export type DocParagraphBlock = {
  /** DocParagraphBlock 在所属文档或任务中的唯一标识。 */
  id: string;
  /** 用于区分 DocParagraphBlock 不同结构分支的类型标识。 */
  type: 'paragraph';
  /** DocParagraphBlock 携带或渲染的文本内容。 */
  text: string;
  /** DocParagraphBlock 包含的 inlines 有序集合。 */
  inlines?: DocTextInline[];
  /** 段落的语义角色，用于选择标题、标题级别或正文样式。 */
  role?: 'title' | 'heading' | 'body';
  /** DocParagraphBlock 使用的渲染或文本样式。 */
  style?: DocTextStyle;
  /** 源文档在该段落前存在显式分页符；渲染分页时该占位段落不显示。 */
  pageBreakBefore?: boolean;
};

/** 描述 DocTableBlock 在 DOC 二进制解析中的数据结构。 */
export type DocTableBlock = {
  /** DocTableBlock 在所属文档或任务中的唯一标识。 */
  id: string;
  /** 用于区分 DocTableBlock 不同结构分支的类型标识。 */
  type: 'table';
  /** DocTableBlock 包含的 rows 有序集合。 */
  rows: DocTableRow[];
  /** DocTableBlock 使用的渲染或文本样式。 */
  style?: DocTableStyle;
  /** 表格各列的标准化宽度，单位为渲染像素。 */
  columns?: number[];
  /** 表格总宽度，单位为标准化渲染像素。 */
  width?: number;
  /** DocTableBlock 的水平对齐方式；未提供时沿用来源格式或渲染器的默认规则。 */
  align?: 'left' | 'center' | 'right';
};

/** 描述 DocTableRow 在 DOC 二进制解析中的数据结构。 */
export type DocTableRow = {
  /** DocTableRow 在所属文档或任务中的唯一标识。 */
  id: string;
  /** DocTableRow 包含的 cells 有序集合。 */
  cells: DocTableCell[];
};

/** 描述 DocTableCell 在 DOC 二进制解析中的数据结构。 */
export type DocTableCell = {
  /** DocTableCell 在所属文档或任务中的唯一标识。 */
  id: string;
  /** DocTableCell 携带或渲染的文本内容。 */
  text: string;
  /** DocTableCell 包含的 inlines 有序集合。 */
  inlines?: DocTextInline[];
  /** DocTableCell 使用的渲染或文本样式。 */
  style?: DocTextStyle;
  /** DocTableCell 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderTop?: string;
  /** DocTableCell 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderRight?: string;
  /** DocTableCell 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderBottom?: string;
  /** DocTableCell 对应方向的 CSS 边框样式；未提供时沿用来源格式或渲染器的默认规则。 */
  borderLeft?: string;
  /** 单元格宽度，单位为标准化渲染像素。 */
  width?: number;
  /** 表格单元格横向跨越的列数。 */
  colSpan?: number;
  /** 表格单元格纵向跨越的行数。 */
  rowSpan?: number;
  /** DocTableCell 的垂直对齐方式；未提供时沿用来源格式或渲染器的默认规则。 */
  verticalAlign?: 'top' | 'middle' | 'bottom';
};

/** 描述 DocListBlock 在 DOC 二进制解析中的数据结构。 */
export type DocListBlock = {
  /** DocListBlock 在所属文档或任务中的唯一标识。 */
  id: string;
  /** 用于区分 DocListBlock 不同结构分支的类型标识。 */
  type: 'list';
  /** 是否按有序列表渲染；false 表示无序列表。 */
  ordered: boolean;
  /** DocListBlock 包含的 items 有序集合。 */
  items: DocListItem[];
  /** DocListBlock 使用的渲染或文本样式。 */
  style?: DocTextStyle;
};

/** 描述 DocListItem 在 DOC 二进制解析中的数据结构。 */
export type DocListItem = {
  /** DocListItem 在所属文档或任务中的唯一标识。 */
  id: string;
  /** DocListItem 携带或渲染的文本内容。 */
  text: string;
  /** DocListItem 包含的 inlines 有序集合。 */
  inlines?: DocTextInline[];
};

/** 描述 DocTextInline 在 DOC 二进制解析中的数据结构。 */
export type DocTextInline = DocTextRunInline | DocImageInline;

/** 描述 DocTextRunInline 在 DOC 二进制解析中的数据结构。 */
export type DocTextRunInline = {
  /** 用于区分 DocTextRunInline 不同结构分支的类型标识。 */
  type: 'text';
  /** DocTextRunInline 携带或渲染的文本内容。 */
  text: string;
  /** DocTextRunInline 使用的渲染或文本样式。 */
  style?: DocTextStyle;
};

/** 描述 DocImageInline 在 DOC 二进制解析中的数据结构。 */
export type DocImageInline = {
  /** 用于区分 DocImageInline 不同结构分支的类型标识。 */
  type: 'image';
  /** DocImageInline 当前关联的图片资源或图片模型。 */
  image: DocImage;
};

/** 描述 DOC 二进制解析使用的样式参数。 */
export type DocTextStyle = {
  /** DocTextStyle 的前景或文本颜色，使用标准化 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
  color?: string;
  /** DocTextStyle 的背景颜色，使用 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
  backgroundColor?: string;
  /** 段落边框颜色；未提供时不绘制边框。 */
  borderColor?: string;
  /** 段落边框宽度，单位为标准化渲染像素。 */
  borderWidth?: number;
  /** 段落边框线型。 */
  borderStyle?: 'solid' | 'dashed' | 'dotted';
  /** DocTextStyle 的字号，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  fontSize?: number;
  /** DocTextStyle 的字体粗细值；未提供时沿用来源格式或渲染器的默认规则。 */
  fontWeight?: number;
  /** DocTextStyle 的字体样式；未提供时沿用来源格式或渲染器的默认规则。 */
  fontStyle?: 'normal' | 'italic';
  /** CSS 文本装饰样式，例如下划线或删除线。 */
  textDecoration?: string;
  /** 段落文本的水平对齐方式。 */
  textAlign?: 'left' | 'center' | 'right' | 'justify';
  /** DocTextStyle 的行高，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  lineHeight?: number;
  /** DocTextStyle 的字体族名称；未提供时沿用来源格式或渲染器的默认规则。 */
  fontFamily?: string;
  /** DocTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  indentLeft?: number;
  /** DocTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  indentRight?: number;
  /** DocTextStyle 的首行缩进，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  firstLineIndent?: number;
  /** DocTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  spacingBefore?: number;
  /** DocTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  spacingAfter?: number;
  /** DocTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingTop?: number;
  /** DocTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingRight?: number;
  /** DocTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingBottom?: number;
  /** DocTextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  paddingLeft?: number;
};

/** 描述 DOC 二进制解析使用的样式参数。 */
export type DocTableStyle = {
  /** 表头单元格使用的背景颜色。 */
  headerBackgroundColor?: string;
  /** DocTableStyle 的 headerTextColor 文本值。 */
  headerTextColor?: string;
  /** DocTableStyle 的 borderColor 文本值。 */
  borderColor?: string;
  /** DocTableStyle 的 cellBackgroundColor 文本值。 */
  cellBackgroundColor?: string;
  /** DocTableStyle 的 stripedRowBackgroundColor 文本值。 */
  stripedRowBackgroundColor?: string;
};

/** 描述 DocImage 在 DOC 二进制解析中的数据结构。 */
export type DocImage = {
  /** DocImage 在所属文档或任务中的唯一标识。 */
  id: string;
  /** DocImage 的 src 文本值。 */
  src: string;
  /** 资源的 MIME 类型，用于选择解码和渲染方式。 */
  mimeType: string;
  /** DocImage 的 width 几何值，单位遵循对应 Office 二进制记录定义；未提供时沿用来源格式或渲染器的默认规则。 */
  width?: number;
  /** DocImage 的 height 几何值，单位遵循对应 Office 二进制记录定义；未提供时沿用来源格式或渲染器的默认规则。 */
  height?: number;
  /** DocImage 的 caption 文本值。 */
  caption?: string;
  /** DocImage 在源二进制流中的字节偏移。 */
  offset?: number;
  /** 图片是否采用锚点定位而非随文排列；未提供时使用来源格式或渲染器的默认行为。 */
  anchored?: boolean;
};
