import type { OfficeHyperlink } from '../../shared/hyperlink';
import type {
  WordInlineReview,
  WordReviewDocument,
} from '../word/review/types';
import type { WordBookmarkTarget, WordOutlineItem } from '../word/types';

/** 包含页面、正文、资源和大纲的标准化 DOC 文档。 */
export type DocDocument = {
  /** 面向用户展示的标题。 */
  title: string;
  /** 当前关联的页面模型。 */
  page: DocPage;
  /** 按源文档顺序排列的内容块。 */
  blocks: DocBlock[];
  /** 按源文档顺序排列的段落。 */
  paragraphs: DocParagraph[];
  /** 当前文档或页面包含的图片资源。 */
  images: DocImage[];
  /** 源 DOC/WPS 明确声明的大纲条目；为空时不显示目录侧栏。 */
  outline?: WordOutlineItem[];
  /** 按源名称索引的文档内部书签。 */
  bookmarks?: Record<string, WordBookmarkTarget>;
  /** 从页眉 story 中识别出的徽标图片；未提供时不渲染页眉图片。 */
  headerImage?: DocImage;
  /** 页脚 story 是否包含 PAGE 字段。 */
  footerPageNumbers?: boolean;
  /** 解析时产生但不阻止继续预览的警告。 */
  warnings: string[];
  /** 源文档中可恢复的批注、修订、脚注和尾注。 */
  review?: WordReviewDocument;
  /** 可按正文引用呈现的脚注和尾注正文。 */
  notes?: DocNotes;
  /** DOC/WPS 文档持有且需要在销毁时释放的浏览器资源。 */
  resources?: DocResources;
};

/** DOC 文档持有且需要统一释放的浏览器资源。 */
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

/** DOC 页面的尺寸和页边距。 */
export type DocPage = {
  /** 标准化页面宽度，单位为渲染像素。 */
  width: number;
  /** 标准化页面最小高度，单位为渲染像素。 */
  minHeight: number;
  /** 上外边距，单位为标准化渲染像素。 */
  marginTop: number;
  /** 右外边距，单位为标准化渲染像素。 */
  marginRight: number;
  /** 下外边距，单位为标准化渲染像素。 */
  marginBottom: number;
  /** 左外边距，单位为标准化渲染像素。 */
  marginLeft: number;
};

/** DOC 段落的纯文本和样式。 */
export type DocParagraph = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 文本内容。 */
  text: string;
};

/** 标准化 DOC 支持的块级内容联合类型。 */
export type DocBlock = DocParagraphBlock | DocTableBlock | DocListBlock;

/** DOC 段落块的文本、行内节点和排版信息。 */
export type DocParagraphBlock = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 分页拆分后仍指向原始正文块的稳定 ID。 */
  sourceBlockId?: string;
  /** 用于区分联合类型分支的类型标识。 */
  type: 'paragraph';
  /** 文本内容。 */
  text: string;
  /** 按源文档顺序排列的行内内容。 */
  inlines?: DocTextInline[];
  /** 段落的语义角色，用于选择标题、标题级别或正文样式。 */
  role?: 'title' | 'heading' | 'body';
  /** 源 DOC/WPS 明确声明的大纲级别，使用从 0 开始的内部表示。 */
  outlineLevel?: number;
  /** 是否属于 Word 自动目录段落。 */
  isTableOfContents?: boolean;
  /** 分页器估算的块高度；仅用于在浏览器布局后校准旧 DOC 分页。 */
  estimatedHeight?: number;
  /** 当前内容使用的渲染样式。 */
  style?: DocTextStyle;
  /** 源文档在该段落前存在显式分页符；渲染分页时该占位段落不显示。 */
  pageBreakBefore?: boolean;
};

/** DOC 表格块的行列、尺寸和排版信息。 */
export type DocTableBlock = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 分页拆分后仍指向原始正文块的稳定 ID。 */
  sourceBlockId?: string;
  /** 用于区分联合类型分支的类型标识。 */
  type: 'table';
  /** 按显示顺序排列的表格行。 */
  rows: DocTableRow[];
  /** 当前内容使用的渲染样式。 */
  style?: DocTableStyle;
  /** 表格各列的标准化宽度，单位为渲染像素。 */
  columns?: number[];
  /** 表格总宽度，单位为标准化渲染像素。 */
  width?: number;
  /** 水平对齐方式。 */
  align?: 'left' | 'center' | 'right';
  /** 表格外边界相对正文左边界的偏移，单位为标准化渲染像素。 */
  offsetLeft?: number;
  /** 表格前的垂直间距；用于保留源文档中位于表格前的空段落。 */
  spacingBefore?: number;
  /** 表格后的垂直间距；用于保留源文档中紧随表格的空段落。 */
  spacingAfter?: number;
};

/** DOC 表格行及其高度规则。 */
export type DocTableRow = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 按显示顺序排列的单元格。 */
  cells: DocTableCell[];
  /** 源 DOC 表格行高度，单位为标准化渲染像素。 */
  height?: number;
  /** 正行高表示最小高度，负行高表示精确高度。 */
  heightRule?: 'atLeast' | 'exact';
};

/** DOC 表格单元格的内容、边框和合并信息。 */
export type DocTableCell = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 文本内容。 */
  text: string;
  /** 按源文档顺序排列的行内内容。 */
  inlines?: DocTextInline[];
  /** 当前内容使用的渲染样式。 */
  style?: DocTextStyle;
  /** 上边框的 CSS 样式。 */
  borderTop?: string;
  /** 右边框的 CSS 样式。 */
  borderRight?: string;
  /** 下边框的 CSS 样式。 */
  borderBottom?: string;
  /** 左边框的 CSS 样式。 */
  borderLeft?: string;
  /** 单元格宽度，单位为标准化渲染像素。 */
  width?: number;
  /** 表格单元格横向跨越的列数。 */
  colSpan?: number;
  /** 表格单元格纵向跨越的行数。 */
  rowSpan?: number;
  /** 垂直对齐方式。 */
  verticalAlign?: 'top' | 'middle' | 'bottom';
};

/** DOC 列表块的编号和列表项。 */
export type DocListBlock = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 分页拆分后仍指向原始正文块的稳定 ID。 */
  sourceBlockId?: string;
  /** 用于区分联合类型分支的类型标识。 */
  type: 'list';
  /** 是否按有序列表渲染；false 表示无序列表。 */
  ordered: boolean;
  /** 按显示顺序排列的项目。 */
  items: DocListItem[];
  /** 当前内容使用的渲染样式。 */
  style?: DocTextStyle;
  /** 当前列表块后续仍有分页片段，页尾不重复追加列表外间距。 */
  continuesOnNext?: boolean;
  /** 分页器估算的块高度；仅用于在浏览器布局后校准旧 DOC 分页。 */
  estimatedHeight?: number;
};

/** DOC 列表中的单个编号项。 */
export type DocListItem = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 源列表显示的编号或项目符号；保留原值才能还原多级编号格式。 */
  marker?: string;
  /** 文本内容。 */
  text: string;
  /** 按源文档顺序排列的行内内容。 */
  inlines?: DocTextInline[];
};

/** DOC 段落支持的文本、图片和书签行内节点。 */
export type DocTextInline =
  | DocTextRunInline
  | DocImageInline
  | DocBookmarkInline
  | DocNoteReferenceInline;

/** 具有统一样式的 DOC 连续文本片段。 */
export type DocTextRunInline = {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'text';
  /** 文本内容。 */
  text: string;
  /** 当前内容使用的渲染样式。 */
  style?: DocTextStyle;
  /** 源文档为该段文字声明的超链接。 */
  hyperlink?: OfficeHyperlink;
  /** 当前文字关联的批注范围和可恢复修订。 */
  review?: WordInlineReview;
};

/** DOC/WPS 脚注或尾注引用的稳定信息。 */
export type DocNoteReference = {
  /** 源注释集合中的稳定标识。 */
  noteId: string;
  /** 脚注显示在引用页，尾注显示在文档末尾。 */
  noteKind: 'footnote' | 'endnote';
  /** 页面中显示的连续引用编号。 */
  label: string;
};

/** DOC/WPS 正文中的脚注或尾注行内引用。 */
export type DocNoteReferenceInline = DocNoteReference & {
  /** 用于区分注释引用和普通文本。 */
  type: 'note-reference';
  /** 引用标记沿用的文字样式。 */
  style?: DocTextStyle;
};

/** 嵌入 DOC 段落中的图片节点。 */
export type DocImageInline = {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'image';
  /** 当前关联的图片资源或图片模型。 */
  image: DocImage;
};

/** DOC 正文中用于内部跳转的零宽书签标记。 */
export type DocBookmarkInline = {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'bookmark';
  /** 源文档声明的书签名称。 */
  name: string;
  /** 渲染定位标记时使用的稳定标识。 */
  markerId: string;
};

/** 单条 DOC/WPS 脚注或尾注正文。 */
export type DocNote = DocNoteReference & {
  /** 注释正文复用标准 DOC 块模型。 */
  blocks: DocBlock[];
};

/** DOC/WPS 文档包含的脚注和尾注集合。 */
export type DocNotes = {
  /** 按源文档顺序排列的脚注。 */
  footnotes: DocNote[];
  /** 按源文档顺序排列的尾注。 */
  endnotes: DocNote[];
};

/** DOC 文本、段落和边框的标准化样式。 */
export type DocTextStyle = {
  /** 前景或文字颜色，使用 CSS 颜色值。 */
  color?: string;
  /** 背景颜色，使用 CSS 颜色值。 */
  backgroundColor?: string;
  /** 标记来自段落属性的底纹，避免字符高亮被错误扩散到整段。 */
  paragraphBackgroundColor?: string;
  /** 段落边框颜色；未提供时不绘制边框。 */
  borderColor?: string;
  /** 段落边框宽度，单位为标准化渲染像素。 */
  borderWidth?: number;
  /** 段落边框线型。 */
  borderStyle?: 'solid' | 'dashed' | 'dotted';
  /** 字号，单位为标准化渲染像素。 */
  fontSize?: number;
  /** 字体粗细。 */
  fontWeight?: number;
  /** 字体样式。 */
  fontStyle?: 'normal' | 'italic';
  /** CSS 文本装饰样式，例如下划线或删除线。 */
  textDecoration?: string;
  /** 段落文本的水平对齐方式。 */
  textAlign?: 'left' | 'center' | 'right' | 'justify';
  /** 行高，单位为标准化渲染像素。 */
  lineHeight?: number;
  /** DOC 的 LSPD 同时声明倍数行距时保留该倍数，供标题和目录布局选用。 */
  lineHeightMultiplier?: number;
  /** 段落是否跟随节内文档网格；false 时使用 Word 自身的自动行距。 */
  useDocumentGrid?: boolean;
  /** 字体族名称。 */
  fontFamily?: string;
  /** 左缩进，单位为标准化渲染像素。 */
  indentLeft?: number;
  /** 右缩进，单位为标准化渲染像素。 */
  indentRight?: number;
  /** 首行缩进，单位为标准化渲染像素。 */
  firstLineIndent?: number;
  /** 段前间距，单位为标准化渲染像素。 */
  spacingBefore?: number;
  /** 段后间距，单位为标准化渲染像素。 */
  spacingAfter?: number;
  /** 上内边距，单位为标准化渲染像素。 */
  paddingTop?: number;
  /** 右内边距，单位为标准化渲染像素。 */
  paddingRight?: number;
  /** 下内边距，单位为标准化渲染像素。 */
  paddingBottom?: number;
  /** 左内边距，单位为标准化渲染像素。 */
  paddingLeft?: number;
};

/** DOC 表格的背景色、交替行和单元格间距样式。 */
export type DocTableStyle = {
  /** 表头单元格使用的背景颜色。 */
  headerBackgroundColor?: string;
  /** 表头 文本 颜色。 */
  headerTextColor?: string;
  /** 边框颜色。 */
  borderColor?: string;
  /** 普通表格单元格使用的背景颜色。 */
  cellBackgroundColor?: string;
  /** 交替表格行使用的背景颜色。 */
  stripedRowBackgroundColor?: string;
};

/** DOC 图片资源的尺寸、位置和环绕信息。 */
export type DocImage = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 图片资源地址或延迟资源引用。 */
  src: string;
  /** 资源的 MIME 类型，用于选择解码和渲染方式。 */
  mimeType: string;
  /** 宽度，单位为标准化渲染像素。 */
  width?: number;
  /** 高度，单位为标准化渲染像素。 */
  height?: number;
  /** 图片对应的说明文字。 */
  caption?: string;
  /** 在所属数据范围中的偏移位置。 */
  offset?: number;
  /** 图片是否采用锚点定位而非随文排列。 */
  anchored?: boolean;
  /** 整页绘图画布相对正文文字所在的叠放层级。 */
  pageDrawingLayer?: 'behindText' | 'inFrontOfText';
  /** 整页锚定画布相对正文区域向页面四边扩展的距离。 */
  pageInsets?: {
    /** 页面正文上边距。 */
    top: number;
    /** 页面正文右边距。 */
    right: number;
    /** 页面正文下边距。 */
    bottom: number;
    /** 页面正文左边距。 */
    left: number;
  };
  /** 源文档为图片声明的超链接。 */
  hyperlink?: OfficeHyperlink;
};
