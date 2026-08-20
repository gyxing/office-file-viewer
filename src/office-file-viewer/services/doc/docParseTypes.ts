import type { ParseProgress } from '../parsing/types';
import type { WordInlineReview } from '../word/review/types';
import type { DocDrawingSection } from './parseDocDrawingCanvas';
import type { DocNumberingCatalog } from './parseDocNumbering';
import type { DocStyleOutlineCatalog } from './parseDocStyleOutline';
import type {
  DocBlock,
  DocImage,
  DocNoteReference,
  DocTableBlock,
  DocTextStyle,
} from './types';

/** DOC 二进制书签表中可以精确映射到正文的书签。 */
export type DocBinaryBookmark = {
  /** 源文档声明的书签名称。 */
  name: string;
  /** 书签起点在主文档 story 中的字符位置。 */
  charStart: number;
  /** 书签终点在主文档 story 中的字符位置。 */
  charEnd: number;
  /** 渲染零宽定位标记时使用的稳定标识。 */
  markerId: string;
};

/** 插入正文片段起点的书签标记。 */
export type DocBookmarkMarker = Pick<DocBinaryBookmark, 'name' | 'markerId'>;

/** DOC Piece Table 中的字符区间和文件偏移。 */
export type DocPiece = {
  /** 对应内容在文档字符流中的起始位置。 */
  charStart: number;
  /** 对应内容在文档字符流中的结束位置。 */
  charEnd: number;
  /** 在对应二进制流中的字节偏移。 */
  fileOffset: number;
  /** 源负载是否使用压缩存储。 */
  compressed: boolean;
};

/** DOC 文件信息块中当前解析链会使用的字段。 */
export type DocFib = {
  /** FIB 标记指定的表流名称。 */
  tableStreamName: '0Table' | '1Table';
  /** 主文档 story 的字符数。 */
  ccpText: number;
  /** 脚注 story 的字符数。 */
  ccpFtn: number;
  /** 页眉页脚 story 的字符数。 */
  ccpHdr: number;
  /** 宏 story 的字符数。 */
  ccpMcr: number;
  /** 批注 story 的字符数。 */
  ccpAtn: number;
  /** 尾注 story 的字符数。 */
  ccpEdn: number;
  /** 文本框 story 的字符数。 */
  ccpTxbx: number;
  /** 样式表在表流中的起始偏移。 */
  fcStshf: number;
  /** 样式表占用的字节数。 */
  lcbStshf: number;
  /** 脚注引用 PLCF 在表流中的起始偏移。 */
  fcPlcffndRef: number;
  /** 脚注引用 PLCF 占用的字节数。 */
  lcbPlcffndRef: number;
  /** 脚注正文边界 PLCF 在表流中的起始偏移。 */
  fcPlcffndTxt: number;
  /** 脚注正文边界 PLCF 占用的字节数。 */
  lcbPlcffndTxt: number;
  /** 批注引用 PLCF 在表流中的起始偏移。 */
  fcPlcfandRef: number;
  /** 批注引用 PLCF 占用的字节数。 */
  lcbPlcfandRef: number;
  /** 批注正文边界 PLCF 在表流中的起始偏移。 */
  fcPlcfandTxt: number;
  /** 批注正文边界 PLCF 占用的字节数。 */
  lcbPlcfandTxt: number;
  /** Section PLC 在表流中的起始偏移。 */
  fcPlcfSed: number;
  /** Section PLC 占用的字节数。 */
  lcbPlcfSed: number;
  /** 字符属性 BTE PLC 的起始偏移。 */
  fcPlcfBteChpx: number;
  /** 字符属性 BTE PLC 占用的字节数。 */
  lcbPlcfBteChpx: number;
  /** 段落属性 BTE PLC 的起始偏移。 */
  fcPlcfBtePapx: number;
  /** 段落属性 BTE PLC 占用的字节数。 */
  lcbPlcfBtePapx: number;
  /** 字体表在表流中的起始偏移。 */
  fcSttbfFfn: number;
  /** 字体表占用的字节数。 */
  lcbSttbfFfn: number;
  /** 标准书签名称表在表流中的起始偏移。 */
  fcSttbfBkmk: number;
  /** 标准书签名称表占用的字节数。 */
  lcbSttbfBkmk: number;
  /** 标准书签起点 PLC 在表流中的起始偏移。 */
  fcPlcfBkf: number;
  /** 标准书签起点 PLC 占用的字节数。 */
  lcbPlcfBkf: number;
  /** 标准书签终点 PLC 在表流中的起始偏移。 */
  fcPlcfBkl: number;
  /** 标准书签终点 PLC 占用的字节数。 */
  lcbPlcfBkl: number;
  /** Piece Table 所属 CLX 的起始偏移。 */
  fcClx: number;
  /** Piece Table 所属 CLX 占用的字节数。 */
  lcbClx: number;
  /** 批注作者名称数组在表流中的起始偏移。 */
  fcGrpXstAtnOwners: number;
  /** 批注作者名称数组占用的字节数。 */
  lcbGrpXstAtnOwners: number;
  /** 主文档绘图锚点 PLC 的起始偏移。 */
  fcPlcSpaMom: number;
  /** 主文档绘图锚点 PLC 占用的字节数。 */
  lcbPlcSpaMom: number;
  /** 批注范围起点 PLC 在表流中的起始偏移。 */
  fcPlcfAtnBkf: number;
  /** 批注范围起点 PLC 占用的字节数。 */
  lcbPlcfAtnBkf: number;
  /** 批注范围终点 PLC 在表流中的起始偏移。 */
  fcPlcfAtnBkl: number;
  /** 批注范围终点 PLC 占用的字节数。 */
  lcbPlcfAtnBkl: number;
  /** 尾注引用 PLCF 在表流中的起始偏移。 */
  fcPlcfendRef: number;
  /** 尾注引用 PLCF 占用的字节数。 */
  lcbPlcfendRef: number;
  /** 尾注正文边界 PLCF 在表流中的起始偏移。 */
  fcPlcfendTxt: number;
  /** 尾注正文边界 PLCF 占用的字节数。 */
  lcbPlcfendTxt: number;
  /** 修订作者字符串表在表流中的起始偏移。 */
  fcSttbfRMark: number;
  /** 修订作者字符串表占用的字节数。 */
  lcbSttbfRMark: number;
  /** OfficeArt 绘图数据的起始偏移。 */
  fcDggInfo: number;
  /** OfficeArt 绘图数据占用的字节数。 */
  lcbDggInfo: number;
  /** 文本框字符范围 PLC 的起始偏移。 */
  fcPlcfTxbxTxt: number;
  /** 文本框字符范围 PLC 占用的字节数。 */
  lcbPlcfTxbxTxt: number;
  /** 列表定义表的起始偏移。 */
  fcPlfLst: number;
  /** 列表定义表占用的字节数。 */
  lcbPlfLst: number;
  /** 列表覆盖表的起始偏移。 */
  fcPlfLfo: number;
  /** 列表覆盖表占用的字节数。 */
  lcbPlfLfo: number;
};

/** DOC 字符属性中可恢复的插入或删除修订。 */
export type DocCharacterRevision = {
  /** 当前字符范围属于插入或删除修订。 */
  kind: 'insert' | 'delete';
  /** 修订作者在 SttbfRMark 中的索引。 */
  authorIndex?: number;
  /** DTTM 解码后的 ISO 本地日期文本。 */
  createdAt?: string;
};

/** DOC 字符属性在文本和文件流中的覆盖范围。 */
export type DocCharacterRun = {
  /** 在 WordDocument 流中的起始字节边界。 */
  fcStart: number;
  /** 在 WordDocument 流中的结束字节边界。 */
  fcEnd: number;
  /** 当前内容使用的渲染样式。 */
  style?: DocTextStyle;
  /** 当前字符范围关联的插入或删除修订。 */
  revisions?: DocCharacterRevision[];
  /** 当前字符范围包含无法还原原始值的属性级修订。 */
  propertyRevision?: boolean;
};

/** TDefTable 中单元格的合并与垂直对齐属性。 */
export type DocTableCellLayout = {
  /** 单元格横向合并的开始或延续状态。 */
  horizontalMerge?: 'continue' | 'restart';
  /** 单元格纵向合并的开始或延续状态。 */
  verticalMerge?: 'continue' | 'restart';
  /** 垂直对齐方式。 */
  verticalAlign?: 'top' | 'middle' | 'bottom';
};

/** DOC 段落属性在文本和文件流中的覆盖范围。 */
export type DocParagraphRun = {
  /** 在 WordDocument 流中的起始字节边界。 */
  fcStart: number;
  /** 在 WordDocument 流中的结束字节边界。 */
  fcEnd: number;
  /** 当前内容使用的渲染样式。 */
  style?: DocTextStyle;
  /** 当前段落或表格属性包含无法还原原始值的修订。 */
  propertyRevision?: boolean;
  /** 当前段落是否位于表格内。 */
  inTable?: boolean;
  /** 当前段落是否为表格行结束标记。 */
  tableRowEnd?: boolean;
  /** 表格行定义提供的高度。 */
  tableRowHeight?: number;
  /** 表格行高度的约束方式。 */
  tableRowHeightRule?: 'atLeast' | 'exact';
  /** 当前段落是否要求在段前强制分页。 */
  pageBreakBefore?: boolean;
  /** 源 PAPX 明确声明的段落大纲级别。 */
  outlineLevel?: number;
  /** 当前段落是否采用自动目录样式。 */
  isTableOfContents?: boolean;
  /** 当前段落引用的一基列表覆盖索引。 */
  listId?: number;
  /** 当前段落使用的零基列表层级。 */
  listLevel?: number;
  /** 表格行定义提供的列宽。 */
  tableColumns?: number[];
  /** 表格行定义提供的水平对齐方式。 */
  tableAlign?: DocTableBlock['align'];
  /** 表格外边界相对正文左边界的偏移。 */
  tableOffsetLeft?: number;
  /** 表格属性声明的首选总宽度。 */
  tableWidth?: number;
  /** 表格行定义提供的逐单元格布局。 */
  tableCellLayouts?: DocTableCellLayout[];
};

/** DOC 正文中的连续文本片段及其样式。 */
export type DocTextSegment = Omit<DocParagraphRun, 'fcStart' | 'fcEnd'> & {
  /** 文本内容。 */
  text: string;
  /** 当前片段在对应 Word story 中的起始字符位置。 */
  charStart?: number;
  /** 当前片段在对应 Word story 中的结束字符位置。 */
  charEnd?: number;
  /** 当前字符位置之前需要渲染的书签标记。 */
  bookmarkMarkers?: DocBookmarkMarker[];
  /** 由 Word 域结果静态解析出的超链接。 */
  hyperlink?: import('../../shared/hyperlink').OfficeHyperlink;
  /** 当前片段对应的脚注或尾注引用。 */
  noteReference?: DocNoteReference;
  /** 当前片段关联的批注范围和可恢复修订。 */
  review?: WordInlineReview;
};

/** DOC 正文中的图片占位片段。 */
export type DocImageSegment = DocTextSegment & {
  /** 当前关联的图片资源或图片模型。 */
  image?: DocImage;
};

/** DOC 节级页面布局及文档网格信息。 */
export type DocSectionLayout = DocDrawingSection & {
  /** 文档网格的行间距，单位为标准化渲染像素。 */
  gridLinePitch?: number;
  /** 当前节属性包含无法还原原始值的修订。 */
  propertyRevision?: boolean;
};

/** DOC 绘图文本框在文本框 story 中的字符区间。 */
export type DocDrawingTextBoxRange = {
  /** 文本框内容的起始字符位置。 */
  charStart: number;
  /** 文本框内容的结束字符位置。 */
  charEnd: number;
};

/** DOC 二进制结构解析后的可复用内容索引。 */
export type DocBinaryContent = {
  /** 当前文档的文件信息块。 */
  fib: DocFib;
  /** 按字符顺序排列的 Piece Table 项。 */
  pieces: DocPiece[];
  /** 按字符范围排列的节级页面布局。 */
  sections: DocSectionLayout[];
  /** 按文件偏移排列的字符样式范围。 */
  characterRuns: DocCharacterRun[];
  /** 按文件偏移排列的段落属性范围。 */
  paragraphRuns: DocParagraphRun[];
  /** 与 OfficeArt ClientTextbox 索引对齐的文本框字符范围。 */
  drawingTextBoxRanges: Array<DocDrawingTextBoxRange | undefined>;
  /** Normal 段落样式继承后的文本样式。 */
  normalStyle?: DocTextStyle;
  /** 段落样式、大纲级别和自动目录规则。 */
  outlineCatalog: DocStyleOutlineCatalog;
  /** 列表定义、覆盖关系和当前计数状态。 */
  numbering: DocNumberingCatalog;
  /** 可精确映射到主文档 story 的标准书签。 */
  bookmarks: DocBinaryBookmark[];
  /** 按索引排列的修订作者名称。 */
  revisionAuthors: string[];
  /** 书签表损坏或暂不支持时产生的降级提示。 */
  bookmarkWarnings: string[];
  /** 文档是否包含尚未支持原始态投影的属性级修订。 */
  hasUnsupportedPropertyRevisions: boolean;
};

/** 将 DOC 文本行组装为块级模型时使用的选项。 */
export type DocBlockBuildOptions = {
  /** 在长任务检查点报告进度并响应取消信号。 */
  checkpoint(progress?: ParseProgress): Promise<void>;
  /** 接收本轮新生成的连续内容块。 */
  onBatch?(startIndex: number, blocks: DocBlock[]): Promise<void>;
  /** 主文档列表格式及当前计数状态。 */
  numbering?: DocNumberingCatalog;
  /** 节属性与默认段落倍数共同计算出的正文网格行高。 */
  defaultGridLineHeight?: number;
  /** 节属性声明的原始文档网格行距。 */
  documentGridLineHeight?: number;
  /** 当前页面扣除页边距后的正文可用高度。 */
  pageContentHeight?: number;
};
