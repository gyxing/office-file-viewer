import { CFB_SIGNATURE, parseCfb } from '../../shared/binary/cfb';
import { createResourceReference } from '../parsing/assembly/resourceReferences';
import type {
  PortableDocMetadata,
  PortableResource,
} from '../parsing/protocol/messages';
import type { ParseProgress } from '../parsing/types';
import {
  documentMetadataFromDoc,
  paragraphsFromDocBlocks,
} from './chunkDocBlocks';
import { DocBlockStreamBuilder } from './DocBlockStreamBuilder';
import {
  extractDocDrawingCanvases,
  type DocDrawingSection,
  type DocDrawingTextBox,
} from './parseDocDrawingCanvas';
import {
  nextDocNumberPrefix,
  readDocNumberingCatalog,
  type DocNumberingCatalog,
} from './parseDocNumbering';
import {
  isDocParagraphTocStyle,
  parseDocStyleOutlineCatalog,
  readDocParagraphOutlineLevel,
  readDocParagraphStyleChain,
  type DocStyleOutlineCatalog,
} from './parseDocStyleOutline';
import type {
  DocBlock,
  DocDocument,
  DocImage,
  DocListBlock,
  DocPage,
  DocParagraphBlock,
  DocTableBlock,
  DocTableStyle,
  DocTextInline,
  DocTextStyle,
} from './types';

/** DOC Piece Table 中的字符区间和文件偏移。 */
type DocPiece = {
  /** 对应内容在文档字符流中的起始位置。 */
  charStart: number;
  /** 对应内容在文档字符流中的结束位置。 */
  charEnd: number;
  /** 在对应二进制流中的字节偏移。 */
  fileOffset: number;
  /** 源负载是否使用压缩存储。 */
  compressed: boolean;
};

/** 解析后的 DOC 文件信息块。 */
type DocFib = ReturnType<typeof parseFib>;

/** DOC 字符属性在文本和文件流中的覆盖范围。 */
type DocCharacterRun = {
  /** 在 WordDocument 流中的字节边界。 */
  fcStart: number;
  /** 在 WordDocument 流中的字节边界。 */
  fcEnd: number;
  /** 当前内容使用的渲染样式。 */
  style: DocTextStyle;
};

/** TDefTable 中单元格的合并与垂直对齐属性。 */
type DocTableCellLayout = {
  /** 单元格横向合并的开始或延续状态。 */
  horizontalMerge?: 'continue' | 'restart';
  /** 单元格纵向合并的开始或延续状态。 */
  verticalMerge?: 'continue' | 'restart';
  /** 垂直对齐方式。 */
  verticalAlign?: 'top' | 'middle' | 'bottom';
};

/** DOC 段落属性在文本和文件流中的覆盖范围。 */
type DocParagraphRun = {
  /** 在 WordDocument 流中的字节边界。 */
  fcStart: number;
  /** 在 WordDocument 流中的字节边界。 */
  fcEnd: number;
  /** 当前内容使用的渲染样式。 */
  style?: DocTextStyle;
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
type DocTextSegment = {
  /** 文本内容。 */
  text: string;
  /** 当前内容使用的渲染样式。 */
  style?: DocTextStyle;
  /** 当前文本片段是否位于表格内。 */
  inTable?: boolean;
  /** 当前文本片段是否结束表格行。 */
  tableRowEnd?: boolean;
  /** 表格行定义提供的高度。 */
  tableRowHeight?: number;
  /** 表格行高度的约束方式。 */
  tableRowHeightRule?: 'atLeast' | 'exact';
  /** 当前文本片段所在段落是否要求段前分页。 */
  pageBreakBefore?: boolean;
  /** 源 PAPX 明确声明的段落大纲级别。 */
  outlineLevel?: number;
  /** 当前文本片段是否采用自动目录样式。 */
  isTableOfContents?: boolean;
  /** 当前文本片段引用的一基列表覆盖索引。 */
  listId?: number;
  /** 当前文本片段使用的零基列表层级。 */
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

/** DOC 正文中的图片占位片段。 */
type DocImageSegment = {
  /** 文本内容。 */
  text: string;
  /** 当前内容使用的渲染样式。 */
  style?: DocTextStyle;
  /** 当前关联的图片资源或图片模型。 */
  image?: DocImage;
  /** 当前文本片段是否位于表格内。 */
  inTable?: boolean;
  /** 当前文本片段是否结束表格行。 */
  tableRowEnd?: boolean;
  /** 表格行定义提供的高度。 */
  tableRowHeight?: number;
  /** 表格行高度的约束方式。 */
  tableRowHeightRule?: 'atLeast' | 'exact';
  /** 当前文本片段所在段落是否要求段前分页。 */
  pageBreakBefore?: boolean;
  /** 源 PAPX 明确声明的段落大纲级别。 */
  outlineLevel?: number;
  /** 当前文本片段是否采用自动目录样式。 */
  isTableOfContents?: boolean;
  /** 当前文本片段引用的一基列表覆盖索引。 */
  listId?: number;
  /** 当前文本片段使用的零基列表层级。 */
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

/** 等待分配资源标识和地址的 DOC 图片。 */
type DocImageCandidate = Omit<DocImage, 'id' | 'src'> & {
  /** 原始字节序列。 */
  bytes: Uint8Array;
  /** 在所属数据范围中的偏移位置。 */
  offset: number;
  /** 占用或消费的字节数。 */
  byteLength: number;
  /** 候选图片是否来自文档内嵌媒体包。 */
  packagedMedia: boolean;
  /** 候选图片是否为 Web 扩展对象的预览图。 */
  webExtensionPreview: boolean;
  /** 图片资源所在的复合文档流名称。 */
  streamName: string;
};

/** 识别为列表项的 DOC 文本行。 */
type ParsedListLine = {
  /** 列表是否使用有序编号。 */
  ordered: boolean;
  /** 文本内容。 */
  text: string;
  /** 按源文档顺序排列的行内内容。 */
  inlines?: DocTextInline[];
  /** 列表段落从源文档继承的字体与行距。 */
  style?: DocTextStyle;
};

/** DOC 页面布局使用的单行文本和行内节点。 */
type DocLine = {
  /** 文本内容。 */
  text: string;
  /** 按源文档顺序排列的行内内容。 */
  inlines: DocTextInline[];
  /** 当前内容使用的渲染样式。 */
  style?: DocTextStyle;
  /** 当前行是否位于表格内。 */
  inTable?: boolean;
  /** 当前行是否为表格行结束位置。 */
  tableRowEnd?: boolean;
  /** 表格行定义提供的高度。 */
  tableRowHeight?: number;
  /** 表格行高度的约束方式。 */
  tableRowHeightRule?: 'atLeast' | 'exact';
  /** 当前行是否要求在段前强制分页。 */
  pageBreakBefore?: boolean;
  /** 源 PAPX 明确声明的段落大纲级别。 */
  outlineLevel?: number;
  /** 当前行是否采用自动目录样式。 */
  isTableOfContents?: boolean;
  /** 当前行引用的一基列表覆盖索引。 */
  listId?: number;
  /** 当前行使用的零基列表层级。 */
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
  /** 执行 操作时调用的函数。 */
  match: (regexp: RegExp) => RegExpMatchArray | null;
};

/** 构建 DOC 表格时尚未完成合并处理的单元格。 */
type PendingTableCell = {
  /** 文本内容。 */
  text: string;
  /** 按源文档顺序排列的行内内容。 */
  inlines: DocTextInline[];
  /** 当前内容使用的渲染样式。 */
  style?: DocTextStyle;
  /** 当前单元格横向跨越的列数。 */
  colSpan?: number;
  /** 当前单元格纵向跨越的行数。 */
  rowSpan?: number;
  /** 当前单元格内容在垂直方向的对齐方式。 */
  verticalAlign?: 'top' | 'middle' | 'bottom';
};

/** 表格行在输出单元格之外保留源 TDefTable 的网格。 */
type PendingTableRow = {
  /** 按显示顺序排列的单元格。 */
  cells: PendingTableCell[];
  /** 按显示顺序排列的列定义。 */
  columns?: number[];
  /** 水平对齐方式。 */
  align?: DocTableBlock['align'];
  /** 表格行相对内容区域左侧的偏移。 */
  offsetLeft?: number;
  /** 宽度，单位为标准化渲染像素。 */
  width?: number;
  /** 当前表格行中各单元格的布局信息。 */
  cellLayouts?: DocTableCellLayout[];
  /** 高度，单位为标准化渲染像素。 */
  height?: number;
  /** 表格行高采用自动、最小值或固定值的规则。 */
  heightRule?: 'atLeast' | 'exact';
};

/** 按字体编号索引的 DOC 字体族名称。 */
type DocFontTable = string[];

/** 汇总 DOC 二进制解析各步骤共享的上下文。 */
export type DocCoreContext = {
  /** 正在解析的原始文件名，用于格式识别和错误提示。 */
  fileName: string;
  /** 在长任务检查点报告进度并响应取消信号。 */
  checkpoint(progress?: ParseProgress): Promise<void>;
  /** 处理完成后生成的输出结果。 */
  output?: DocCoreOutput;
};

/** DOC 核心解析生成的文档及性能档案。 */
export type DocCoreOutput = {
  /** 接收解析器产生的可移植资源分块。 */
  resource(resource: PortableResource): Promise<void>;
  /** 接收文字文档的主体元数据。 */
  documentMetadata(metadata: PortableDocMetadata): Promise<void>;
  /** 接收文字文档的连续内容块。 */
  documentBlocks(startIndex: number, blocks: DocBlock[]): Promise<void>;
};

/** DOC 核心解析成功或失败的联合结果。 */
export type DocCoreResult = {
  /** 当前处理的标准化文档模型。 */
  document: DocDocument;
  /** 持有的图片、字体或对象 URL 等资源；文档释放时需同步清理。 */
  resources: PortableResource[];
};

/** 提供已经由随机 CFB Reader 读取的 DOC 核心流。 */
export type DocCoreStreamsInput = {
  /** WordDocument 主流。 */
  wordDocument: Uint8Array;
  /** FIB 指定的 0Table 或 1Table 流。 */
  tableStream: Uint8Array;
  /** 仅用于提取图片资源的相关流，通常为 WordDocument、Table 和 Data。 */
  imageStreams: Iterable<readonly [string, Uint8Array]>;
};

/** 将 DOC 文本行组装为块级模型时使用的选项。 */
type DocBlockBuildOptions = {
  /** 在长任务检查点报告进度并响应取消信号。 */
  checkpoint(progress?: ParseProgress): Promise<void>;
  /** 接收本轮新生成的连续内容块。 */
  onBatch?(startIndex: number, blocks: DocBlock[]): Promise<void>;
  /** 主文档列表格式及当前计数状态。 */
  numbering?: DocNumberingCatalog;
  /** 节属性与默认段落倍数共同计算出的正文网格行高。 */
  defaultGridLineHeight?: number;
  /** 节属性声明的原始文档网格行距，供表格内紧凑文本保持源行距。 */
  documentGridLineHeight?: number;
  /** 当前页面扣除页边距后的正文可用高度，用于长目录单页收敛。 */
  pageContentHeight?: number;
};

// 旧版 .doc 是 OLE/CFB 二进制容器，不是 zip；这里实现最小可用的前端降级解析。
/** DOC 缺少页面设置时使用的默认页面尺寸和边距。 */
const DEFAULT_DOC_PAGE = {
  width: 794,
  minHeight: 1123,
  marginTop: 96,
  marginRight: 120,
  marginBottom: 96,
  marginLeft: 120,
};

/** DOC 无法解析字体信息时使用的默认字体回退栈。 */
const DOC_FONT_FAMILY =
  '"Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", Arial, sans-serif';
/** Word 旧版颜色索引到 CSS 颜色值的映射。 */
const WORD_ICO_COLORS: Record<number, string> = {
  1: '#000000',
  2: '#0000ff',
  3: '#00ffff',
  4: '#00ff00',
  5: '#ff00ff',
  6: '#ff0000',
  7: '#ffff00',
  8: '#ffffff',
  9: '#000080',
  10: '#008080',
  11: '#008000',
  12: '#800080',
  13: '#800000',
  14: '#808000',
  15: '#808080',
  16: '#c0c0c0',
};

function isOleDoc(bytes: Uint8Array) {
  return CFB_SIGNATURE.every((value, index) => bytes[index] === value);
}

function readUint16(view: DataView, offset: number) {
  return view.getUint16(offset, true);
}

function readUint32(view: DataView, offset: number) {
  return view.getUint32(offset, true);
}

function readUint16BE(view: DataView, offset: number) {
  return view.getUint16(offset, false);
}

function readUint32BE(view: DataView, offset: number) {
  return view.getUint32(offset, false);
}

function readInt16(view: DataView, offset: number) {
  return view.getInt16(offset, true);
}

function twipToPx(value: number) {
  return (value / 1440) * 96;
}

function readFibField(wordDocument: Uint8Array, offset: number) {
  if (offset + 4 > wordDocument.length) return 0;
  return readUint32(
    new DataView(
      wordDocument.buffer,
      wordDocument.byteOffset,
      wordDocument.byteLength,
    ),
    offset,
  );
}

function parseFib(wordDocument: Uint8Array) {
  const view = new DataView(
    wordDocument.buffer,
    wordDocument.byteOffset,
    wordDocument.byteLength,
  );
  const flags = readUint16(view, 10);

  return {
    tableStreamName: flags & 0x0200 ? '1Table' : '0Table',
    ccpText: readFibField(wordDocument, 76),
    ccpFtn: readFibField(wordDocument, 80),
    ccpHdr: readFibField(wordDocument, 84),
    ccpMcr: readFibField(wordDocument, 88),
    ccpAtn: readFibField(wordDocument, 92),
    ccpEdn: readFibField(wordDocument, 96),
    ccpTxbx: readFibField(wordDocument, 100),
    fcStshf: readFibField(wordDocument, 154),
    lcbStshf: readFibField(wordDocument, 158),

    fcPlcfSed: readFibField(wordDocument, 202),
    lcbPlcfSed: readFibField(wordDocument, 206),
    fcPlcfBteChpx: readFibField(wordDocument, 250),
    lcbPlcfBteChpx: readFibField(wordDocument, 254),
    fcPlcfBtePapx: readFibField(wordDocument, 258),
    lcbPlcfBtePapx: readFibField(wordDocument, 262),
    fcSttbfFfn: readFibField(wordDocument, 274),
    lcbSttbfFfn: readFibField(wordDocument, 278),
    fcClx: readFibField(wordDocument, 418),
    lcbClx: readFibField(wordDocument, 422),
    fcPlcSpaMom: readFibField(wordDocument, 474),
    lcbPlcSpaMom: readFibField(wordDocument, 478),
    fcDggInfo: readFibField(wordDocument, 554),
    lcbDggInfo: readFibField(wordDocument, 558),
    fcPlcfTxbxTxt: readFibField(wordDocument, 602),
    lcbPlcfTxbxTxt: readFibField(wordDocument, 606),
    fcPlfLst: readFibField(wordDocument, 738),
    lcbPlfLst: readFibField(wordDocument, 742),
    fcPlfLfo: readFibField(wordDocument, 746),
    lcbPlfLfo: readFibField(wordDocument, 750),
  };
}

/** 查找 `findPieceTable` 对应的目标数据。 */
function findPieceTable(clx: Uint8Array) {
  let offset = 0;
  const view = new DataView(clx.buffer, clx.byteOffset, clx.byteLength);

  while (offset < clx.length) {
    const type = clx[offset];

    if (type === 0x02) {
      const length = readUint32(view, offset + 1);
      return clx.slice(offset + 5, offset + 5 + length);
    }

    if (type === 0x01) {
      const length = readUint16(view, offset + 1);
      offset += 3 + length;
      continue;
    }

    offset += 1;
  }

  return undefined;
}

function parsePieces(tableStream: Uint8Array, fib: DocFib) {
  // Piece table 描述正文字符区间与 WordDocument 字节偏移的映射，是读取 DOC 正文的核心索引。
  const clx = tableStream.slice(fib.fcClx, fib.fcClx + fib.lcbClx);
  const pieceTable = findPieceTable(clx);
  if (!pieceTable) return [];

  const pieceCount = Math.floor((pieceTable.length - 4) / 12);
  const view = new DataView(
    pieceTable.buffer,
    pieceTable.byteOffset,
    pieceTable.byteLength,
  );
  const pieces: DocPiece[] = [];

  for (let index = 0; index < pieceCount; index += 1) {
    const charStart = readUint32(view, index * 4);
    const charEnd = readUint32(view, (index + 1) * 4);
    const pcdOffset = (pieceCount + 1) * 4 + index * 8;
    const fcValue = readUint32(view, pcdOffset + 2);
    const compressed = Boolean(fcValue & 0x40000000);
    const fileOffset = compressed ? (fcValue & 0x3fffffff) / 2 : fcValue;

    if (charEnd > charStart && fileOffset >= 0) {
      pieces.push({ charStart, charEnd, fileOffset, compressed });
    }
  }

  return pieces;
}

/** 按文档字符区间裁剪 piece table，并同步修正底层流偏移。 */
function slicePiecesByCharacterRange(
  pieces: DocPiece[],
  rangeStart: number,
  rangeEnd: number,
) {
  if (rangeEnd <= rangeStart) return [];
  return pieces.flatMap((piece) => {
    const charStart = Math.max(piece.charStart, rangeStart);
    const charEnd = Math.min(piece.charEnd, rangeEnd);
    if (charEnd <= charStart) return [];
    const localStart = charStart - piece.charStart;
    return [
      {
        ...piece,
        charStart,
        charEnd,
        fileOffset:
          piece.fileOffset + (piece.compressed ? localStart : localStart * 2),
      },
    ];
  });
}

/** DOC 常见中文字体名称到可用字体族的回退映射。 */
const DOC_FONT_ALIASES: Record<string, string[]> = {
  宋体: ['SimSun'],
  新宋体: ['NSimSun'],
  黑体: ['SimHei'],
  楷体: ['KaiTi'],
  楷体_GB2312: ['KaiTi_GB2312', 'KaiTi'],
  仿宋: ['FangSong'],
  仿宋_GB2312: ['FangSong_GB2312', 'FangSong'],
};

/** 把 Word 字体名称转换为兼容中英文系统名称的 CSS 字体栈。 */
function quoteFontFamily(value: string | undefined) {
  if (!value) return undefined;
  const fonts = value
    .split(',')
    .map((font) => font.trim().replace(/^['"]|['"]$/g, ''))
    .filter(Boolean)
    .flatMap((font) => [...(DOC_FONT_ALIASES[font] ?? []), font]);
  return [...new Set(fonts)].map((font) => `"${font}"`).join(', ');
}

/** 把西文字体放在东亚字体之前，交由浏览器按字符缺字规则选择实际字形。 */
function appendFontFamilyFallback(
  primary: string | undefined,
  fallback: string | undefined,
) {
  if (!primary) return fallback;
  if (!fallback) return primary;
  const primaryFonts = new Set(
    primary
      .split(',')
      .map((font) => font.replace(/["']/g, '').trim().toLowerCase()),
  );
  const missingFallbacks = fallback
    .split(',')
    .map((font) => font.trim())
    .filter(
      (font) =>
        !primaryFonts.has(font.replace(/["']/g, '').trim().toLowerCase()),
    );
  return missingFallbacks.length
    ? `${primary}, ${missingFallbacks.join(', ')}`
    : primary;
}

function parseFontTable(tableStream: Uint8Array, fib: DocFib): DocFontTable {
  if (!fib.fcSttbfFfn || !fib.lcbSttbfFfn) return [];
  const data = tableStream.slice(
    fib.fcSttbfFfn,
    fib.fcSttbfFfn + fib.lcbSttbfFfn,
  );
  if (data.length < 4) return [];

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const extended = readUint16(view, 0) === 0xffff;
  const count =
    extended && data.length >= 6 ? readUint16(view, 2) : readUint16(view, 0);
  const fonts: string[] = [];
  let offset = extended ? 6 : 4;

  for (let index = 0; index < count && offset + 1 < data.length; index += 1) {
    const entryLength = (data[offset] ?? 0) + 1;
    if (entryLength <= 1 || offset + entryLength > data.length) break;

    const nameOffset = offset + 40;
    if (nameOffset < offset + entryLength) {
      const rawName = new TextDecoder('utf-16le').decode(
        data.slice(nameOffset, offset + entryLength),
      );
      const name = rawName
        .split('\u0000')[0]
        ?.replace(/\uFFFD/g, '')
        .trim();
      if (name) fonts.push(name);
    }

    offset += entryLength;
  }

  return fonts;
}

function plcItemCount(length: number, dataSize: number) {
  return Math.floor((length - 4) / (4 + dataSize));
}

function parsePlcBteChpx(tableStream: Uint8Array, fib: DocFib) {
  if (!fib.fcPlcfBteChpx || !fib.lcbPlcfBteChpx) return [];

  const data = tableStream.slice(
    fib.fcPlcfBteChpx,
    fib.fcPlcfBteChpx + fib.lcbPlcfBteChpx,
  );
  const count = plcItemCount(data.length, 4);
  if (count <= 0) return [];

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const pnOffset = (count + 1) * 4;

  return Array.from({ length: count }, (_, index) => ({
    fcStart: readUint32(view, index * 4),
    fcEnd: readUint32(view, (index + 1) * 4),
    pn: readUint32(view, pnOffset + index * 4) & 0x003fffff,
  })).filter((item) => item.fcEnd > item.fcStart);
}

function parsePlcBtePapx(tableStream: Uint8Array, fib: DocFib) {
  if (!fib.fcPlcfBtePapx || !fib.lcbPlcfBtePapx) return [];

  const data = tableStream.slice(
    fib.fcPlcfBtePapx,
    fib.fcPlcfBtePapx + fib.lcbPlcfBtePapx,
  );
  const count = plcItemCount(data.length, 4);
  if (count <= 0) return [];

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const pnOffset = (count + 1) * 4;

  return Array.from({ length: count }, (_, index) => ({
    fcStart: readUint32(view, index * 4),
    fcEnd: readUint32(view, (index + 1) * 4),
    pn: readUint32(view, pnOffset + index * 4) & 0x003fffff,
  })).filter((item) => item.fcEnd > item.fcStart);
}

/** 合并 `mergeTextStyle` 接收的多份数据。 */
function mergeTextStyle(
  base: DocTextStyle | undefined,
  next: DocTextStyle | undefined,
): DocTextStyle | undefined {
  if (!base && !next) return undefined;
  return {
    ...base,
    ...next,
  };
}

/** 合并 `mergeStyleIntoTextStyle` 接收的多份数据。 */
function mergeStyleIntoTextStyle(
  base: DocTextStyle,
  override: DocTextStyle | undefined,
) {
  return {
    ...base,
    ...override,
    fontSize: override?.fontSize ?? base.fontSize,
    fontWeight: override?.fontWeight ?? base.fontWeight,
    fontStyle: override?.fontStyle ?? base.fontStyle,
    textDecoration: override?.textDecoration ?? base.textDecoration,
    color: override?.color ?? base.color,
    backgroundColor: override?.backgroundColor ?? base.backgroundColor,
    textAlign: override?.textAlign ?? base.textAlign,
    lineHeight: override?.lineHeight ?? base.lineHeight,
    useDocumentGrid: override?.useDocumentGrid ?? base.useDocumentGrid,
    fontFamily: override?.fontFamily ?? base.fontFamily,
    indentLeft: override?.indentLeft ?? base.indentLeft,
    indentRight: override?.indentRight ?? base.indentRight,
    firstLineIndent: override?.firstLineIndent ?? base.firstLineIndent,
    spacingBefore: override?.spacingBefore ?? base.spacingBefore,
    spacingAfter: override?.spacingAfter ?? base.spacingAfter,
    paddingTop: override?.paddingTop ?? base.paddingTop,
    paddingRight: override?.paddingRight ?? base.paddingRight,
    paddingBottom: override?.paddingBottom ?? base.paddingBottom,
    paddingLeft: override?.paddingLeft ?? base.paddingLeft,
  };
}

function tableCellTextStyle(
  style: DocTextStyle | undefined,
): DocTextStyle | undefined {
  if (!style) return undefined;
  const cellTextStyle: DocTextStyle = {
    color: style.color,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight,
    fontStyle: style.fontStyle,
    textDecoration: style.textDecoration,
    textAlign: style.textAlign,
    lineHeight: style.lineHeight,
    lineHeightMultiplier: style.lineHeightMultiplier,
    useDocumentGrid: style.useDocumentGrid,
    fontFamily: style.fontFamily,
  };
  const cleaned = Object.fromEntries(
    Object.entries(cellTextStyle).filter(([, value]) => value !== undefined),
  ) as DocTextStyle;
  return Object.keys(cleaned).length ? cleaned : undefined;
}

/** 合并 `mergeTextDecoration` 接收的多份数据。 */
function mergeTextDecoration(
  style: DocTextStyle,
  decoration: string,
  enabled: boolean,
) {
  const values = new Set(
    (style.textDecoration ?? '').split(/\s+/).filter(Boolean),
  );
  if (enabled) values.add(decoration);
  else values.delete(decoration);
  style.textDecoration = values.size ? [...values].join(' ') : undefined;
}

/** 把 DOC 的 COLORREF（低字节依次为 RGB）转换为 CSS 颜色。 */
function readDocColorRef(operand: Uint8Array, offset: number) {
  if (offset < 0 || offset + 4 > operand.length) return undefined;
  const red = operand[offset] ?? 0;
  const green = operand[offset + 1] ?? 0;
  const blue = operand[offset + 2] ?? 0;
  const flags = operand[offset + 3] ?? 0;
  // 0xFFFFFFFF 表示自动颜色，应继续沿用文档或渲染器的默认值。
  if (red === 0xff && green === 0xff && blue === 0xff && flags === 0xff) {
    return undefined;
  }
  return `#${[red, green, blue]
    .map((value) => value.toString(16).padStart(2, '0'))
    .join('')}`;
}

/** 解析新版段落边框 BrcOperand，并同步恢复边框与文字间距。 */
function applyDocParagraphBorder(
  style: DocTextStyle,
  sprm: number,
  operand: Uint8Array,
) {
  // 可变长操作数首字节为 cb，后续 8 字节才是实际 Brc。
  if (operand.length < 9 || (operand[0] ?? 0) < 8) return;
  const color = readDocColorRef(operand, 1);
  const widthEighthPoints = operand[5] ?? 0;
  const borderType = operand[6] ?? 0;
  const flags = readUint16(
    new DataView(operand.buffer, operand.byteOffset, operand.byteLength),
    7,
  );
  const distancePoints = flags & 0x1f;

  if (borderType !== 0xff && widthEighthPoints > 0) {
    style.borderColor = color ?? style.borderColor ?? '#000000';
    style.borderWidth = (widthEighthPoints / 8) * (96 / 72);
    style.borderStyle =
      borderType === 6
        ? 'dotted'
        : borderType === 7 || borderType === 8
        ? 'dashed'
        : 'solid';
  }

  // Word 分边记录文字到边框的距离；保留各边距离可避免内容紧贴边线。
  const padding = distancePoints * (96 / 72);
  if (sprm === 0xc64e) style.paddingTop = padding;
  if (sprm === 0xc64f) style.paddingLeft = padding;
  if (sprm === 0xc650) style.paddingBottom = padding;
  if (sprm === 0xc651) style.paddingRight = padding;
}

/** 把 `applySprmOperand` 对应的规则应用到目标对象。 */
function applySprmOperand(
  style: DocTextStyle,
  sprm: number,
  operand: Uint8Array,
  fonts: DocFontTable = [],
) {
  const operandView = new DataView(
    operand.buffer,
    operand.byteOffset,
    operand.byteLength,
  );
  const first = operand[0];

  if ((sprm === 0x0835 || sprm === 0x0800) && first !== undefined) {
    style.fontWeight = first ? 700 : 400;
    return;
  }

  if ((sprm === 0x0836 || sprm === 0x0801) && first !== undefined) {
    style.fontStyle = first ? 'italic' : 'normal';
    return;
  }

  if ((sprm === 0x0837 || sprm === 0x0802) && first !== undefined) {
    mergeTextDecoration(style, 'underline', first !== 0);
    return;
  }

  if ((sprm === 0x0838 || sprm === 0x0803) && first !== undefined) {
    mergeTextDecoration(style, 'line-through', first !== 0);
    return;
  }

  if ((sprm === 0x4a43 || sprm === 0x4a4d) && operand.length >= 2) {
    const halfPoints = readInt16(operandView, 0);
    if (halfPoints > 0 && halfPoints < 200) {
      style.fontSize = (halfPoints / 2) * (96 / 72);
    }
    return;
  }

  if (
    (sprm === 0x4a4f ||
      sprm === 0x4a50 ||
      sprm === 0x4a51 ||
      sprm === 0x4a4e) &&
    operand.length >= 2
  ) {
    const font = quoteFontFamily(fonts[readUint16(operandView, 0)]);
    if (font && sprm === 0x4a50) {
      // Word 同一字符样式可同时声明西文与东亚字体，CSS 字体栈必须保留两者。
      style.fontFamily = appendFontFamilyFallback(style.fontFamily, font);
    } else if (font && !style.fontFamily) {
      style.fontFamily = font;
    }
    return;
  }

  if ((sprm === 0x2a42 || sprm === 0x2a24) && first !== undefined) {
    style.color = WORD_ICO_COLORS[first];
    return;
  }

  if (sprm === 0x2a0c && first !== undefined) {
    // sprmCHighlight 使用 Word 的 ico 调色板，0 表示清除字符高亮。
    style.backgroundColor = first === 0 ? undefined : WORD_ICO_COLORS[first];
    return;
  }

  if (sprm === 0x442d && operand.length >= 2) {
    const shading = readUint16(operandView, 0);
    const backgroundColor = WORD_ICO_COLORS[(shading >> 5) & 0x1f];
    if (backgroundColor) {
      style.backgroundColor = backgroundColor;
      style.paragraphBackgroundColor = backgroundColor;
    }
    return;
  }

  if (sprm === 0xc64d && operand.length >= 11) {
    // ShdOperand 首字节为 cb，cvBack 从实际 Shd 的第 5 字节开始。
    const backgroundColor = readDocColorRef(operand, 5);
    if (backgroundColor) {
      style.backgroundColor = backgroundColor;
      style.paragraphBackgroundColor = backgroundColor;
    }
    return;
  }

  if (
    sprm === 0xc64e ||
    sprm === 0xc64f ||
    sprm === 0xc650 ||
    sprm === 0xc651
  ) {
    applyDocParagraphBorder(style, sprm, operand);
    return;
  }

  if (sprm === 0x2447 && first !== undefined) {
    // sprmPFUsePgsuSettings 为 0 时，Word 表格段落不跟随节内文档网格。
    style.useDocumentGrid = first !== 0;
    return;
  }

  if ((sprm === 0x2403 || sprm === 0x2461) && first !== undefined) {
    const alignment = ['left', 'center', 'right', 'justify'][first];
    if (alignment) style.textAlign = alignment as DocTextStyle['textAlign'];
    return;
  }

  if ((sprm === 0x840f || sprm === 0x845e) && operand.length >= 2) {
    style.indentLeft = twipToPx(readInt16(operandView, 0));
    return;
  }

  if ((sprm === 0x840e || sprm === 0x845d) && operand.length >= 2) {
    style.indentRight = twipToPx(readInt16(operandView, 0));
    return;
  }

  if ((sprm === 0x8411 || sprm === 0x8460) && operand.length >= 2) {
    style.firstLineIndent = twipToPx(readInt16(operandView, 0));
    return;
  }

  if ((sprm === 0xa413 || sprm === 0xa416) && operand.length >= 2) {
    const value = readInt16(operandView, 0);
    if (value >= 0) style.spacingBefore = twipToPx(value);
    return;
  }

  if ((sprm === 0xa414 || sprm === 0xa417) && operand.length >= 2) {
    const value = readInt16(operandView, 0);
    if (value >= 0) style.spacingAfter = twipToPx(value);
    return;
  }

  if ((sprm === 0x6412 || sprm === 0x6461) && operand.length >= 4) {
    const line = readInt16(operandView, 0);
    const usesMultiplier = readUint16(operandView, 2) === 1;
    if (line !== 0) {
      // LSPD 明确区分倍数行距与绝对 twip 行距，CSS 需沿用相同语义。
      style.lineHeight = usesMultiplier
        ? Math.abs(line) / 240
        : twipToPx(Math.abs(line));
      style.lineHeightMultiplier = usesMultiplier
        ? Math.abs(line) / 240
        : undefined;
    }
  }
}

function sprmOperandSize(sprm: number, bytes: Uint8Array, offset: number) {
  if (sprm === 0xd608 && offset + 2 <= bytes.length) {
    // TDefTableOperand 使用 16 位 cb；其值等于操作数总长度减一。
    return (
      readUint16(
        new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength),
        offset,
      ) + 1
    );
  }
  const sizeCode = (sprm >> 13) & 0x7;
  if (sizeCode === 0 || sizeCode === 1) return 1;
  if (sizeCode === 2 || sizeCode === 4 || sizeCode === 5) return 2;
  if (sizeCode === 3) return 4;
  if (sizeCode === 6) {
    const length = bytes[offset] ?? 0;
    return 1 + length;
  }
  if (sizeCode === 7) return 3;
  return 0;
}

/** 拼接展开后的段落属性，避免巨大 PAPX 丢失表格行定义。 */
function concatDocPropertyChunks(chunks: Uint8Array[]) {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const output = new Uint8Array(total);
  let offset = 0;
  chunks.forEach((chunk) => {
    output.set(chunk, offset);
    offset += chunk.length;
  });
  return output;
}

/** 从 Data 流读取 sprmPHugePapx 或 sprmPTableProps 指向的 PrcData。 */
function readReferencedParagraphProperties(
  dataStream: Uint8Array | undefined,
  offset: number,
) {
  if (!dataStream || offset < 0 || offset + 2 > dataStream.length) {
    return undefined;
  }
  const view = new DataView(
    dataStream.buffer,
    dataStream.byteOffset,
    dataStream.byteLength,
  );
  const length = readInt16(view, offset);
  if (length <= 0 || offset + 2 + length > dataStream.length) return undefined;
  return dataStream.slice(offset + 2, offset + 2 + length);
}

/** 递归展开 Data 流中的巨大 PAPX 与共用表格属性，并保留有效的行级直接覆盖。 */
function expandParagraphProperties(
  grpprl: Uint8Array,
  dataStream?: Uint8Array,
  visited = new Set<number>(),
  depth = 0,
): Uint8Array {
  if (!dataStream || !grpprl.length || depth >= 8) return grpprl;
  const view = new DataView(
    grpprl.buffer,
    grpprl.byteOffset,
    grpprl.byteLength,
  );
  const chunks: Uint8Array[] = [];
  let offset = 0;
  while (offset + 2 <= grpprl.length) {
    const entryStart = offset;
    const sprm = readUint16(view, offset);
    offset += 2;
    const operandSize = sprmOperandSize(sprm, grpprl, offset);
    if (!operandSize || offset + operandSize > grpprl.length) break;
    const entryEnd = offset + operandSize;
    const isHugePapx = sprm === 0x6646 && entryStart === 0;
    const isTableProperties = sprm === 0x646b;
    if ((isHugePapx || isTableProperties) && operandSize >= 4) {
      const dataOffset = readUint32(view, offset);
      if (!visited.has(dataOffset)) {
        const referenced = readReferencedParagraphProperties(
          dataStream,
          dataOffset,
        );
        if (referenced) {
          visited.add(dataOffset);
          chunks.push(
            expandParagraphProperties(
              referenced,
              dataStream,
              visited,
              depth + 1,
            ),
          );
          if (isHugePapx) return concatDocPropertyChunks(chunks);
          // sprmPTableProps 只把共用表格属性放进 Data 流；行级 TDefTable 等覆盖项仍紧随引用。
          // 继续读取剩余 PAPX，才能恢复真实列宽、合并关系和行高。
          offset = entryEnd;
          continue;
        }
      }
    }
    chunks.push(grpprl.slice(entryStart, entryEnd));
    offset = entryEnd;
  }
  return concatDocPropertyChunks(chunks);
}

/** DOC 节级页面布局及文档网格信息。 */
type DocSectionLayout = DocDrawingSection & {
  /** 文档网格的行间距，单位为标准化渲染像素。 */
  gridLinePitch?: number;
};

/** 从 PlcfSed 指向的 Sepx 中读取各节页面尺寸、页边距和文档网格。 */
function readDocSections(
  wordDocument: Uint8Array,
  tableStream: Uint8Array,
  fib: DocFib,
): DocSectionLayout[] {
  if (!fib.fcPlcfSed || !fib.lcbPlcfSed) return [];
  if (
    fib.fcPlcfSed < 0 ||
    fib.fcPlcfSed + fib.lcbPlcfSed > tableStream.length
  ) {
    return [];
  }

  const sectionCount = plcItemCount(fib.lcbPlcfSed, 12);
  if (sectionCount <= 0) return [];
  const tableView = new DataView(
    tableStream.buffer,
    tableStream.byteOffset,
    tableStream.byteLength,
  );
  const wordView = new DataView(
    wordDocument.buffer,
    wordDocument.byteOffset,
    wordDocument.byteLength,
  );
  const sedOffset = fib.fcPlcfSed + (sectionCount + 1) * 4;
  const sections: DocSectionLayout[] = [];

  for (let index = 0; index < sectionCount; index += 1) {
    const page: DocPage = {
      ...(sections[sections.length - 1]?.page ?? DEFAULT_DOC_PAGE),
    };
    let gridMode = 0;
    let linePitch: number | undefined;
    const fcSepx = tableView.getInt32(sedOffset + index * 12 + 2, true);
    if (fcSepx >= 0 && fcSepx + 2 <= wordDocument.length) {
      const length = readInt16(wordView, fcSepx);
      if (length > 0 && fcSepx + 2 + length <= wordDocument.length) {
        const grpprl = wordDocument.slice(fcSepx + 2, fcSepx + 2 + length);
        const grpprlView = new DataView(
          grpprl.buffer,
          grpprl.byteOffset,
          grpprl.byteLength,
        );
        let offset = 0;
        while (offset + 2 <= grpprl.length) {
          const sprm = readUint16(grpprlView, offset);
          offset += 2;
          const operandSize = sprmOperandSize(sprm, grpprl, offset);
          if (!operandSize || offset + operandSize > grpprl.length) break;
          if (operandSize >= 2) {
            const value = readInt16(grpprlView, offset);
            if (sprm === 0xb01f && value > 0) page.width = twipToPx(value);
            if (sprm === 0xb020 && value > 0) {
              page.minHeight = twipToPx(value);
            }
            if (sprm === 0xb021) page.marginLeft = twipToPx(Math.abs(value));
            if (sprm === 0xb022) page.marginRight = twipToPx(Math.abs(value));
            if (sprm === 0x9023) page.marginTop = twipToPx(Math.abs(value));
            if (sprm === 0x9024) {
              page.marginBottom = twipToPx(Math.abs(value));
            }
            if (sprm === 0x5032) gridMode = readUint16(grpprlView, offset);
            if (sprm === 0x9031) linePitch = readUint16(grpprlView, offset);
          }
          offset += operandSize;
        }
      }
    }

    if (
      page.marginLeft + page.marginRight >= page.width ||
      page.marginTop + page.marginBottom >= page.minHeight
    ) {
      Object.assign(page, DEFAULT_DOC_PAGE);
    }
    sections.push({
      charStart: tableView.getUint32(fib.fcPlcfSed + index * 4, true),
      charEnd: tableView.getUint32(fib.fcPlcfSed + (index + 1) * 4, true),
      page,
      gridLinePitch:
        gridMode !== 0 && linePitch !== undefined && linePitch > 0
          ? twipToPx(linePitch)
          : undefined,
    });
  }

  return sections;
}

/** DOC 绘图文本框在正文字符流中的范围。 */
type DrawingTextBoxRange = {
  /** 对应内容在文档字符流中的起始位置。 */
  charStart: number;
  /** 对应内容在文档字符流中的结束位置。 */
  charEnd: number;
};

/** 从 PlcftxbxTxt 读取每个 OfficeArt 文本框在文本框 story 中的字符范围。 */
function readDrawingTextBoxRanges(
  tableStream: Uint8Array,
  fib: DocFib,
): Array<DrawingTextBoxRange | undefined> {
  if (!fib.fcPlcfTxbxTxt || !fib.lcbPlcfTxbxTxt) return [];
  if (
    fib.fcPlcfTxbxTxt < 0 ||
    fib.fcPlcfTxbxTxt + fib.lcbPlcfTxbxTxt > tableStream.length
  ) {
    return [];
  }
  const count = plcItemCount(fib.lcbPlcfTxbxTxt, 22);
  if (count <= 0) return [];
  const view = new DataView(
    tableStream.buffer,
    tableStream.byteOffset,
    tableStream.byteLength,
  );
  const recordOffset = fib.fcPlcfTxbxTxt + (count + 1) * 4;
  return Array.from(
    { length: count },
    (_, index): DrawingTextBoxRange | undefined => {
      const shapeId = view.getUint32(recordOffset + index * 22 + 14, true);
      const charStart = view.getUint32(fib.fcPlcfTxbxTxt + index * 4, true);
      const charEnd = view.getUint32(fib.fcPlcfTxbxTxt + (index + 1) * 4, true);
      return shapeId && charEnd > charStart && charStart < fib.ccpTxbx
        ? {
            charStart,
            charEnd: Math.min(charEnd, fib.ccpTxbx),
          }
        : undefined;
    },
  );
}
/** 从 TDefTableOperand 读取当前表格行的逻辑列宽。 */
function readTableDefinitionColumns(operand: Uint8Array) {
  if (operand.length < 5) return undefined;
  const columnCount = operand[2] ?? 0;
  const boundaryCount = columnCount + 1;
  if (!columnCount || 3 + boundaryCount * 2 > operand.length) return undefined;
  const view = new DataView(
    operand.buffer,
    operand.byteOffset,
    operand.byteLength,
  );
  const boundaries = Array.from({ length: boundaryCount }, (_, index) =>
    readInt16(view, 3 + index * 2),
  );
  const columns = boundaries
    .slice(0, -1)
    .map((left, index) => twipToPx(boundaries[index + 1] - left));
  if (!columns.every((width) => width >= 0)) return undefined;

  const cellLayoutOffset = 3 + boundaryCount * 2;
  const cellLayouts = Array.from(
    {
      length: Math.min(
        columnCount,
        Math.floor((operand.length - cellLayoutOffset) / 20),
      ),
    },
    (_, index): DocTableCellLayout => {
      const flags = readUint16(view, cellLayoutOffset + index * 20);
      const horizontalMerge = flags & 0x03;
      const verticalMerge = (flags >> 5) & 0x03;
      const verticalAlign = (flags >> 7) & 0x03;
      return {
        horizontalMerge:
          horizontalMerge === 1
            ? 'continue'
            : horizontalMerge >= 2
            ? 'restart'
            : undefined,
        verticalMerge:
          verticalMerge === 1
            ? 'continue'
            : verticalMerge === 3
            ? 'restart'
            : undefined,
        verticalAlign:
          verticalAlign === 1
            ? 'middle'
            : verticalAlign === 2
            ? 'bottom'
            : 'top',
      };
    },
  );
  return {
    columns,
    cellLayouts,
    offsetLeft: twipToPx(boundaries[0]),
  };
}

function parseGrpprlStyle(
  bytes: Uint8Array,
  fonts: DocFontTable = [],
): DocTextStyle | undefined {
  const style: DocTextStyle = {};
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = 0;

  while (offset + 2 <= bytes.length) {
    const sprm = readUint16(view, offset);
    offset += 2;
    const operandSize = sprmOperandSize(sprm, bytes, offset);
    if (!operandSize || offset + operandSize > bytes.length) break;
    applySprmOperand(
      style,
      sprm,
      bytes.slice(offset, offset + operandSize),
      fonts,
    );
    offset += operandSize;
  }

  return Object.keys(style).length ? style : undefined;
}

/** 从 PAPX 属性中读取二进制 DOC 的段落结构标志。 */
function parseGrpprlParagraphStructure(
  bytes: Uint8Array,
  includesStyleIndex = true,
) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // PAPX 的 grpprl 前两个字节是段落样式索引 istd，实际 SPRM 从其后开始。
  let offset = includesStyleIndex ? Math.min(2, bytes.length) : 0;
  let inTable: boolean | undefined;
  let tableRowEnd: boolean | undefined;
  let tableRowHeight: number | undefined;
  let tableRowHeightRule: 'atLeast' | 'exact' | undefined;
  let pageBreakBefore: boolean | undefined;
  let listId: number | undefined;
  let listLevel: number | undefined;
  let tableColumns: number[] | undefined;
  let tableAlign: DocTableBlock['align'];
  let tableCellLayouts: DocTableCellLayout[] | undefined;
  let tableDefinitionOffsetLeft: number | undefined;
  let tableLeft: number | undefined;
  let tableGapHalf: number | undefined;
  let tableWidthUnit: number | undefined;
  let tableWidthValue: number | undefined;

  while (offset + 2 <= bytes.length) {
    const sprm = readUint16(view, offset);
    offset += 2;
    const operandSize = sprmOperandSize(sprm, bytes, offset);
    if (!operandSize || offset + operandSize > bytes.length) break;
    const enabled = bytes[offset] !== 0;
    if (sprm === 0x2416 || sprm === 0x244b) inTable = enabled;
    if (sprm === 0x2417 || sprm === 0x244c) tableRowEnd = enabled;
    if (sprm === 0x9407 && operandSize >= 2) {
      const rawHeight = readInt16(view, offset);
      if (rawHeight !== 0) {
        tableRowHeight = twipToPx(Math.abs(rawHeight));
        tableRowHeightRule = rawHeight < 0 ? 'exact' : 'atLeast';
      }
    }
    if (sprm === 0x2407) pageBreakBefore = enabled;
    if (sprm === 0x260a) listLevel = bytes[offset];
    if (sprm === 0x460b && operandSize >= 2) listId = readInt16(view, offset);
    if (sprm === 0xd608) {
      const definition = readTableDefinitionColumns(
        bytes.slice(offset, offset + operandSize),
      );
      tableColumns = definition?.columns;
      tableCellLayouts = definition?.cellLayouts;
      tableDefinitionOffsetLeft = definition?.offsetLeft;
    }
    if (sprm === 0x5400 && operandSize >= 2) {
      const value = readUint16(view, offset);
      tableAlign = value === 1 ? 'center' : value === 2 ? 'right' : 'left';
    }
    if (sprm === 0x9601 && operandSize >= 2) {
      tableLeft = readInt16(view, offset);
    }
    if (sprm === 0x9602 && operandSize >= 2) {
      tableGapHalf = readInt16(view, offset);
    }
    if (sprm === 0xf614 && operandSize >= 3) {
      tableWidthUnit = bytes[offset];
      tableWidthValue = readUint16(view, offset + 1);
    }
    offset += operandSize;
  }

  const tableOffsetLeft =
    tableLeft !== undefined || tableGapHalf !== undefined
      ? twipToPx((tableLeft ?? 0) - (tableGapHalf ?? 0))
      : tableDefinitionOffsetLeft;
  const tableContentWidth =
    DEFAULT_DOC_PAGE.width -
    DEFAULT_DOC_PAGE.marginLeft -
    DEFAULT_DOC_PAGE.marginRight;
  const tableWidth =
    tableWidthUnit === 2 && tableWidthValue !== undefined
      ? (tableContentWidth * tableWidthValue) / 5000 +
        Math.max(0, -(tableOffsetLeft ?? 0)) * 2
      : tableWidthUnit === 3 && tableWidthValue !== undefined
      ? twipToPx(tableWidthValue)
      : undefined;

  return {
    inTable,
    tableRowEnd,
    tableRowHeight,
    tableRowHeightRule,
    pageBreakBefore,
    listId,
    listLevel,
    tableColumns,
    tableAlign,
    tableOffsetLeft,
    tableWidth,
    tableCellLayouts,
  };
}

/** DOC 段落在字符流中的边界和结构信息。 */
type DocParagraphStructure = ReturnType<typeof parseGrpprlParagraphStructure>;

/** 仅用后续明确声明的属性覆盖继承结构，避免 undefined 清除父样式。 */
function mergeParagraphStructure(
  base: DocParagraphStructure,
  next: DocParagraphStructure,
) {
  const merged = { ...base };
  Object.entries(next).forEach(([key, value]) => {
    if (value !== undefined) {
      merged[key as keyof DocParagraphStructure] = value as never;
    }
  });
  return merged;
}

function parseChpxFkpPage(
  wordDocument: Uint8Array,
  pageOffset: number,
  fonts: DocFontTable,
): DocCharacterRun[] {
  const page = wordDocument.slice(pageOffset, pageOffset + 512);
  if (page.length < 512) return [];

  const runCount = page[511] ?? 0;
  const view = new DataView(page.buffer, page.byteOffset, page.byteLength);
  const runs: DocCharacterRun[] = [];

  for (let index = 0; index < runCount; index += 1) {
    const fcStart = readUint32(view, index * 4);
    const fcEnd = readUint32(view, (index + 1) * 4);
    const chpxOffset = page[(runCount + 1) * 4 + index];
    if (!chpxOffset || fcEnd <= fcStart) continue;

    const chpxStart = chpxOffset * 2;
    const chpxLength = page[chpxStart] ?? 0;
    const grpprl = page.slice(chpxStart + 1, chpxStart + 1 + chpxLength);
    const style = parseGrpprlStyle(grpprl, fonts);
    if (style) runs.push({ fcStart, fcEnd, style });
  }

  return runs;
}

function parseCharacterRuns(
  wordDocument: Uint8Array,
  tableStream: Uint8Array,
  fib: DocFib,
  fonts: DocFontTable,
): DocCharacterRun[] {
  return parsePlcBteChpx(tableStream, fib).flatMap((entry) =>
    parseChpxFkpPage(wordDocument, entry.pn * 512, fonts).filter(
      (run) => run.fcEnd > entry.fcStart && run.fcStart < entry.fcEnd,
    ),
  );
}

/** 合并 STSH 段落样式继承链中的 PAPX/CHPX 格式差异。 */
function readInheritedParagraphStyle(
  grpprl: Uint8Array,
  catalog: DocStyleOutlineCatalog,
  fonts: DocFontTable,
) {
  return readDocParagraphStyleChain(grpprl, catalog).reduce<
    DocTextStyle | undefined
  >(
    (style, definition) => {
      const paragraphStyle = definition.paragraphGrpprl
        ? parseGrpprlStyle(definition.paragraphGrpprl, fonts)
        : undefined;
      const characterStyle = definition.characterGrpprl
        ? parseGrpprlStyle(definition.characterGrpprl, fonts)
        : undefined;
      return mergeTextStyle(
        mergeTextStyle(style, paragraphStyle),
        characterStyle,
      );
    },
    { fontWeight: 400, fontStyle: 'normal' },
  );
}

/** 合并 STSH 段落样式继承链中的列表和表格结构属性。 */
function readInheritedParagraphStructure(
  grpprl: Uint8Array,
  catalog: DocStyleOutlineCatalog,
) {
  return readDocParagraphStyleChain(grpprl, catalog).reduce(
    (structure, definition) =>
      definition.paragraphGrpprl
        ? mergeParagraphStructure(
            structure,
            parseGrpprlParagraphStructure(definition.paragraphGrpprl, false),
          )
        : structure,
    {} as DocParagraphStructure,
  );
}

function parsePapxFkpPage(
  wordDocument: Uint8Array,
  pageOffset: number,
  outlineCatalog: DocStyleOutlineCatalog,
  fonts: DocFontTable,
  dataStream?: Uint8Array,
): DocParagraphRun[] {
  const page = wordDocument.slice(pageOffset, pageOffset + 512);
  if (page.length < 512) return [];

  const runCount = page[511] ?? 0;
  const view = new DataView(page.buffer, page.byteOffset, page.byteLength);
  const runs: DocParagraphRun[] = [];

  for (let index = 0; index < runCount; index += 1) {
    const fcStart = readUint32(view, index * 4);
    const fcEnd = readUint32(view, (index + 1) * 4);
    const bxOffset = (runCount + 1) * 4 + index * 13;
    const papxOffset = page[bxOffset];
    if (!papxOffset || fcEnd <= fcStart) continue;

    const papxStart = papxOffset * 2;
    const cb = page[papxStart] ?? 0;
    const cbPrime = cb === 0 ? page[papxStart + 1] ?? 0 : cb;
    const papxLength = cb === 0 ? cbPrime * 2 : cb * 2 - 1;
    const grpprlStart = papxStart + (cb === 0 ? 2 : 1);
    const grpprl = page.slice(grpprlStart, grpprlStart + papxLength);
    // PAPX 的 grpprl 以两字节 istd 开头；巨大 PAPX 的直接属性存放在 Data 流中。
    const stylePrefix = grpprl.slice(0, Math.min(2, grpprl.length));
    const directGrpprl = expandParagraphProperties(
      grpprl.slice(stylePrefix.length),
      dataStream,
    );
    const resolvedGrpprl = concatDocPropertyChunks([stylePrefix, directGrpprl]);
    const style = mergeTextStyle(
      readInheritedParagraphStyle(grpprl, outlineCatalog, fonts),
      parseGrpprlStyle(directGrpprl, fonts),
    );
    const structure = mergeParagraphStructure(
      readInheritedParagraphStructure(grpprl, outlineCatalog),
      parseGrpprlParagraphStructure(directGrpprl, false),
    );
    // 表格单元格可能复用标题样式，但 Word 导航窗格只读取正文段落。
    const outlineLevel = structure.inTable
      ? undefined
      : readDocParagraphOutlineLevel(resolvedGrpprl, outlineCatalog);
    const isTableOfContents = isDocParagraphTocStyle(
      resolvedGrpprl,
      outlineCatalog,
    );
    if (
      style ||
      structure.inTable !== undefined ||
      structure.tableRowEnd !== undefined ||
      structure.tableRowHeight !== undefined ||
      structure.pageBreakBefore !== undefined ||
      structure.listId !== undefined ||
      structure.listLevel !== undefined ||
      structure.tableColumns !== undefined ||
      structure.tableCellLayouts !== undefined ||
      structure.tableAlign !== undefined ||
      structure.tableOffsetLeft !== undefined ||
      structure.tableWidth !== undefined ||
      outlineLevel !== undefined ||
      isTableOfContents
    ) {
      runs.push({
        fcStart,
        fcEnd,
        style,
        ...structure,
        outlineLevel,
        isTableOfContents,
      });
    }
  }

  return runs;
}

function parseParagraphRuns(
  wordDocument: Uint8Array,
  tableStream: Uint8Array,
  fib: DocFib,
  outlineCatalog: DocStyleOutlineCatalog,
  fonts: DocFontTable,
  dataStream?: Uint8Array,
): DocParagraphRun[] {
  return parsePlcBtePapx(tableStream, fib).flatMap((entry) =>
    parsePapxFkpPage(
      wordDocument,
      entry.pn * 512,
      outlineCatalog,
      fonts,
      dataStream,
    ).filter((run) => run.fcEnd > entry.fcStart && run.fcStart < entry.fcEnd),
  );
}

function fileOffsetForPieceChar(piece: DocPiece, charOffset: number) {
  return piece.compressed
    ? (piece.fileOffset + charOffset) * 2
    : piece.fileOffset + charOffset * 2;
}

function pieceCharOffsetForFileOffset(piece: DocPiece, fileOffset: number) {
  return piece.compressed
    ? fileOffset / 2 - piece.fileOffset
    : (fileOffset - piece.fileOffset) / 2;
}

function styleForRange(
  byteStart: number,
  byteEnd: number,
  characterRuns: DocCharacterRun[],
) {
  return characterRuns
    .filter((run) => run.fcEnd > byteStart && run.fcStart < byteEnd)
    .sort((left, right) => left.fcStart - right.fcStart)
    .reduce<DocTextStyle | undefined>(
      (style, run) => mergeTextStyle(style, run.style),
      undefined,
    );
}

function paragraphStyleForRange(
  byteStart: number,
  byteEnd: number,
  paragraphRuns: DocParagraphRun[],
) {
  return paragraphRuns
    .filter((run) => run.fcEnd > byteStart && run.fcStart < byteEnd)
    .sort((left, right) => left.fcStart - right.fcStart)
    .reduce<DocTextStyle | undefined>(
      (style, run) => mergeTextStyle(style, run.style),
      undefined,
    );
}

/** 读取指定字节范围对应的表格段落结构。 */
function paragraphStructureForRange(
  byteStart: number,
  byteEnd: number,
  paragraphRuns: DocParagraphRun[],
) {
  return paragraphRuns
    .filter((run) => run.fcEnd > byteStart && run.fcStart < byteEnd)
    .sort((left, right) => left.fcStart - right.fcStart)
    .reduce(
      (structure, run) => ({
        inTable: run.inTable ?? structure.inTable,
        tableRowEnd: run.tableRowEnd ?? structure.tableRowEnd,
        tableRowHeight: run.tableRowHeight ?? structure.tableRowHeight,
        tableRowHeightRule:
          run.tableRowHeightRule ?? structure.tableRowHeightRule,
        pageBreakBefore: run.pageBreakBefore ?? structure.pageBreakBefore,
        outlineLevel: run.outlineLevel ?? structure.outlineLevel,
        isTableOfContents: run.isTableOfContents || structure.isTableOfContents,
        listId: run.listId ?? structure.listId,
        listLevel: run.listLevel ?? structure.listLevel,
        tableColumns: run.tableColumns ?? structure.tableColumns,
        tableAlign: run.tableAlign ?? structure.tableAlign,
        tableOffsetLeft: run.tableOffsetLeft ?? structure.tableOffsetLeft,
        tableWidth: run.tableWidth ?? structure.tableWidth,
        tableCellLayouts: run.tableCellLayouts ?? structure.tableCellLayouts,
      }),
      {
        inTable: undefined as boolean | undefined,
        tableRowEnd: undefined as boolean | undefined,
        tableRowHeight: undefined as number | undefined,
        tableRowHeightRule: undefined as 'atLeast' | 'exact' | undefined,
        pageBreakBefore: undefined as boolean | undefined,
        outlineLevel: undefined as number | undefined,
        isTableOfContents: false,
        listId: undefined as number | undefined,
        listLevel: undefined as number | undefined,
        tableColumns: undefined as number[] | undefined,
        tableAlign: undefined as DocTableBlock['align'],
        tableOffsetLeft: undefined as number | undefined,
        tableWidth: undefined as number | undefined,
        tableCellLayouts: undefined as DocTableCellLayout[] | undefined,
      },
    );
}

/** 按 `splitPieceByStyleRuns` 的规则拆分输入数据。 */
function splitPieceByStyleRuns(
  piece: DocPiece,
  characterRuns: DocCharacterRun[],
  paragraphRuns: DocParagraphRun[],
) {
  const charLength = piece.charEnd - piece.charStart;
  const byteStart = piece.compressed ? piece.fileOffset * 2 : piece.fileOffset;
  const byteEnd = byteStart + charLength * 2;
  const boundaries = new Set([0, charLength]);

  [...characterRuns, ...paragraphRuns].forEach((run) => {
    if (run.fcEnd <= byteStart || run.fcStart >= byteEnd) return;
    const start = Math.max(
      0,
      Math.floor(
        pieceCharOffsetForFileOffset(piece, Math.max(run.fcStart, byteStart)),
      ),
    );
    const end = Math.min(
      charLength,
      Math.ceil(
        pieceCharOffsetForFileOffset(piece, Math.min(run.fcEnd, byteEnd)),
      ),
    );
    boundaries.add(start);
    boundaries.add(end);
  });

  const sorted = [...boundaries].sort((left, right) => left - right);
  return sorted
    .slice(0, -1)
    .map((start, index) => ({ start, end: sorted[index + 1] }))
    .filter((range) => range.end > range.start);
}

function textSegmentsFromPieces(
  wordDocument: Uint8Array,
  pieces: DocPiece[],
  characterRuns: DocCharacterRun[],
  paragraphRuns: DocParagraphRun[],
) {
  return pieces.flatMap((piece) => {
    return splitPieceByStyleRuns(piece, characterRuns, paragraphRuns).map(
      (range) => {
        const scopedPiece: DocPiece = {
          ...piece,
          charStart: piece.charStart + range.start,
          charEnd: piece.charStart + range.end,
          fileOffset: piece.compressed
            ? piece.fileOffset + range.start
            : piece.fileOffset + range.start * 2,
        };
        const byteStart = fileOffsetForPieceChar(piece, range.start);
        const byteEnd = fileOffsetForPieceChar(piece, range.end);
        const structure = paragraphStructureForRange(
          byteStart,
          byteEnd,
          paragraphRuns,
        );
        return {
          ...readPieceSegment(
            wordDocument,
            scopedPiece,
            mergeTextStyle(
              paragraphStyleForRange(byteStart, byteEnd, paragraphRuns),
              styleForRange(byteStart, byteEnd, characterRuns),
            ),
          ),
          ...structure,
        };
      },
    );
  });
}

/** 解码 `decodeCodePage1252` 接收的源数据。 */
function decodeCodePage1252(bytes: Uint8Array) {
  const decoder =
    typeof TextDecoder !== 'undefined'
      ? new TextDecoder('windows-1252')
      : undefined;
  if (decoder) return decoder.decode(bytes);
  return Array.from(bytes, (value) => String.fromCharCode(value)).join('');
}

function scoreDecodedText(text: string) {
  let score = 0;

  for (const char of text) {
    const code = char.codePointAt(0) ?? 0;
    if (char === '\uFFFD') {
      score -= 12;
    } else if (code === 0) {
      score -= 5;
    } else if (code < 32 && code !== 9 && code !== 10 && code !== 13) {
      score -= 3;
    } else if (/[A-Za-z0-9]/.test(char)) {
      score += 1.2;
    } else if (/\p{Script=Han}/u.test(char)) {
      score += 3.5;
    } else if (/\s/.test(char)) {
      score += 0.2;
    } else if (/[.,:;!?()\-_/|]/.test(char)) {
      score += 0.5;
    } else {
      score += 0.3;
    }
  }

  return score;
}

/** 解码 `decodeCompressedPiece` 接收的源数据。 */
function decodeCompressedPiece(bytes: Uint8Array) {
  const candidates = ['gb18030', 'gbk', 'utf-8', 'windows-1252'];
  let best = decodeCodePage1252(bytes);
  let bestScore = scoreDecodedText(best);

  for (const encoding of candidates) {
    try {
      const decoder = new TextDecoder(encoding, { fatal: false });
      const decoded = decoder.decode(bytes);
      const score = scoreDecodedText(decoded);
      if (score > bestScore) {
        best = decoded;
        bestScore = score;
      }
    } catch {
      // Ignore unsupported encodings and keep the best available decode.
    }
  }

  return best;
}

function readPieceText(wordDocument: Uint8Array, piece: DocPiece) {
  const charLength = piece.charEnd - piece.charStart;
  if (piece.compressed) {
    return decodeCompressedPiece(
      wordDocument.slice(piece.fileOffset, piece.fileOffset + charLength),
    );
  }

  const byteLength = charLength * 2;
  const bytes = wordDocument.slice(
    piece.fileOffset,
    piece.fileOffset + byteLength,
  );
  return new TextDecoder('utf-16le').decode(bytes);
}

function readPieceSegment(
  wordDocument: Uint8Array,
  piece: DocPiece,
  style?: DocTextStyle,
): DocTextSegment {
  return {
    text: readPieceText(wordDocument, piece),
    style,
  };
}

/** DOC 域代码或域结果对应的文本片段。 */
type DocFieldTextChunk = {
  /** 文本内容。 */
  text: string;
  /** 当前文本块对应的源文档片段。 */
  segment: DocTextSegment;
};

/** 跨样式片段保留 Word 字段结果，避免字段指令被 run 边界拆开后泄露为正文。 */
function preserveDocFieldResultSegments(segments: DocTextSegment[]) {
  type FieldFrame = {
    instruction: DocFieldTextChunk[];
    result: DocFieldTextChunk[];
    inResult: boolean;
  };

  const frames: FieldFrame[] = [];
  const output: DocFieldTextChunk[] = [];
  const appendTo = (
    target: DocFieldTextChunk[],
    chunks: DocFieldTextChunk[],
  ) => {
    chunks.forEach((chunk) => {
      if (!chunk.text) return;
      const previous = target[target.length - 1];
      if (previous?.segment === chunk.segment) {
        previous.text += chunk.text;
      } else {
        target.push({ ...chunk });
      }
    });
  };
  const append = (chunks: DocFieldTextChunk[]) => {
    if (!chunks.length) return;
    const frame = frames[frames.length - 1];
    if (!frame) {
      appendTo(output, chunks);
    } else if (frame.inResult) {
      appendTo(frame.result, chunks);
    } else {
      appendTo(frame.instruction, chunks);
    }
  };
  const visibleFieldValue = (frame: FieldFrame) =>
    frame.result.length
      ? frame.result
      : frame.instruction.flatMap((chunk) => {
          const anchors = chunk.text.match(/[\u0001\u0008]/g)?.join('') ?? '';
          return anchors ? [{ ...chunk, text: anchors }] : [];
        });

  segments.forEach((segment) => {
    let textStart = 0;
    for (let index = 0; index < segment.text.length; index += 1) {
      const character = segment.text[index];
      if (
        character !== '\u0013' &&
        character !== '\u0014' &&
        character !== '\u0015'
      ) {
        continue;
      }
      append([{ text: segment.text.slice(textStart, index), segment }]);
      if (character === '\u0013') {
        frames.push({ instruction: [], result: [], inResult: false });
      } else if (character === '\u0014') {
        const frame = frames[frames.length - 1];
        if (frame) frame.inResult = true;
      } else if (character === '\u0015') {
        const frame = frames.pop();
        if (frame) append(visibleFieldValue(frame));
      }
      textStart = index + 1;
    }
    append([{ text: segment.text.slice(textStart), segment }]);
  });

  while (frames.length) {
    const frame = frames.pop();
    if (frame) append(visibleFieldValue(frame));
  }

  return output.map(({ text, segment }) => ({ ...segment, text }));
}

/** 将输入标准化为 `normalizeDocText` 返回的结构。 */
function normalizeDocText(text: string) {
  return text
    .replace(/\u0000/g, '')
    .replace(/\u0007/g, '|')
    .replace(/\u000b/g, '\n')
    .replace(/\u000d/g, '\n')
    .replace(/[\u0002-\u0006\u000e-\u001f]/g, '');
}

/** 将输入标准化为 `normalizeDocTextSegments` 返回的结构。 */
function normalizeDocTextSegments(
  segments: DocTextSegment[],
  images: DocImage[] = [],
  drawingImages: Array<DocImage | undefined> = [],
) {
  let imageIndex = 0;
  let drawingImageIndex = 0;

  return preserveDocFieldResultSegments(segments).flatMap((segment) => {
    const normalizedText = normalizeDocText(segment.text);

    return normalizedText
      .split(/(\n|\f|\u0001|\u0008)/)
      .map((text): DocImageSegment => {
        if (text === '\u0001') {
          const image = images[imageIndex];
          if (image) imageIndex += 1;
          return {
            text,
            style: segment.style,
            image,
            inTable: segment.inTable,
            tableRowEnd: segment.tableRowEnd,
            tableRowHeight: segment.tableRowHeight,
            tableRowHeightRule: segment.tableRowHeightRule,
            pageBreakBefore: segment.pageBreakBefore,
            outlineLevel: segment.outlineLevel,
            isTableOfContents: segment.isTableOfContents,
            listId: segment.listId,
            listLevel: segment.listLevel,
            tableColumns: segment.tableColumns,
            tableAlign: segment.tableAlign,
            tableOffsetLeft: segment.tableOffsetLeft,
            tableWidth: segment.tableWidth,
            tableCellLayouts: segment.tableCellLayouts,
          };
        }
        if (text === '\u0008') {
          const image = drawingImages[drawingImageIndex];
          // 每个绘图标记都消费一个槽位；组合画布后的空槽不能阻塞后续绘图映射。
          drawingImageIndex += 1;
          return {
            text,
            style: segment.style,
            image,
            inTable: segment.inTable,
            tableRowEnd: segment.tableRowEnd,
            tableRowHeight: segment.tableRowHeight,
            tableRowHeightRule: segment.tableRowHeightRule,
            pageBreakBefore: segment.pageBreakBefore,
            outlineLevel: segment.outlineLevel,
            isTableOfContents: segment.isTableOfContents,
            listId: segment.listId,
            listLevel: segment.listLevel,
            tableColumns: segment.tableColumns,
            tableAlign: segment.tableAlign,
            tableOffsetLeft: segment.tableOffsetLeft,
            tableWidth: segment.tableWidth,
            tableCellLayouts: segment.tableCellLayouts,
          };
        }
        return {
          text,
          style: segment.style,
          inTable: segment.inTable,
          tableRowEnd: segment.tableRowEnd,
          tableRowHeight: segment.tableRowHeight,
          tableRowHeightRule: segment.tableRowHeightRule,
          pageBreakBefore: segment.pageBreakBefore,
          outlineLevel: segment.outlineLevel,
          isTableOfContents: segment.isTableOfContents,
          listId: segment.listId,
          listLevel: segment.listLevel,
          tableColumns: segment.tableColumns,
          tableAlign: segment.tableAlign,
          tableOffsetLeft: segment.tableOffsetLeft,
          tableWidth: segment.tableWidth,
          tableCellLayouts: segment.tableCellLayouts,
        };
      })
      .filter(
        (item) =>
          item.image ||
          (item.text.length &&
            item.text !== '\u0001' &&
            item.text !== '\u0008'),
      );
  });
}

/** 将输入标准化为 `normalizeBlockText` 返回的结构。 */
function normalizeBlockText(text: string) {
  return text.replace(/[ \t]+/g, ' ').trim();
}

function textFromInlines(inlines: DocTextInline[]) {
  return inlines
    .map((inline) => (inline.type === 'text' ? inline.text : ''))
    .join('');
}

function sameInlineStyle(left?: DocTextStyle, right?: DocTextStyle) {
  return JSON.stringify(left ?? {}) === JSON.stringify(right ?? {});
}

/** 合并 `mergeAdjacentInlines` 接收的多份数据。 */
function mergeAdjacentInlines(inlines: DocTextInline[]) {
  const merged: DocTextInline[] = [];

  inlines.forEach((inline) => {
    if (inline.type === 'image') {
      merged.push({ ...inline });
      return;
    }
    if (!inline.text) return;
    const previous = merged[merged.length - 1];
    if (
      previous?.type === 'text' &&
      sameInlineStyle(previous.style, inline.style)
    ) {
      previous.text += inline.text;
      return;
    }
    merged.push({ ...inline });
  });

  return merged;
}

function trimInlines(inlines: DocTextInline[]) {
  const result = inlines
    .map((inline) => ({ ...inline }))
    .filter(
      (inline) =>
        inline.type === 'image' ||
        (inline.type === 'text' && inline.text.length),
    );

  while (result.length && result[0].type === 'text' && !result[0].text.trim()) {
    result.shift();
  }

  while (result.length) {
    const last = result[result.length - 1];
    if (last.type !== 'text' || last.text.trim()) break;
    result.pop();
  }

  if (result.length) {
    if (result[0].type === 'text')
      result[0].text = result[0].text.replace(/^\s+/, '');
    const last = result[result.length - 1];
    if (last.type === 'text') last.text = last.text.replace(/\s+$/, '');
  }

  return mergeAdjacentInlines(result);
}

function looksLikeTableRow(line: string) {
  return (
    line
      .split('|')
      .map((cell) => cell.trim())
      .filter(Boolean).length >= 2
  );
}

/** 按 `splitTableCells` 的规则拆分输入数据。 */
function splitTableCells(line: DocLine): PendingTableCell[] {
  const cells: PendingTableCell[] = [];
  let current: DocTextInline[] = [];
  const textInlines = (inlines: DocTextInline[]) =>
    inlines.filter(
      (
        item,
      ): item is Extract<
        DocTextInline,
        {
          /** 固定为 `text`，用于区分联合类型分支。 */
          type: 'text';
        }
      > => item.type === 'text',
    );

  line.inlines.forEach((inline) => {
    if (inline.type === 'image') {
      current.push(inline);
      return;
    }
    const parts = inline.text.split('|');

    parts.forEach((part, index) => {
      if (index > 0) {
        const inlines = trimInlines(current);
        cells.push({
          text: normalizeBlockText(textFromInlines(inlines)),
          inlines,
          style:
            dominantStyle(
              textInlines(inlines).map((item) => ({
                text: item.text,
                style: item.style,
              })),
            ) ?? line.style,
        });
        current = [];
      }

      if (part) {
        current.push({ ...inline, text: part });
      }
    });
  });

  const inlines = trimInlines(current);
  if (inlines.length) {
    cells.push({
      text: normalizeBlockText(textFromInlines(inlines)),
      inlines,
      style: dominantStyle(
        textInlines(inlines).map((item) => ({
          text: item.text,
          style: item.style,
        })),
      ),
    });
  }

  return cells;
}

function sliceLineInlines(line: DocLine, start: number) {
  let offset = 0;
  const result: DocTextInline[] = [];

  line.inlines.forEach((inline) => {
    if (inline.type === 'image') return;
    const inlineStart = offset;
    const inlineEnd = inlineStart + inline.text.length;
    offset = inlineEnd;

    if (inlineEnd <= start) return;

    result.push({
      ...inline,
      text: inline.text.slice(Math.max(0, start - inlineStart)),
    });
  });

  return trimInlines(result);
}

function parseListLine(line: DocLine): ParsedListLine | undefined {
  const orderedMatch = line.text.match(
    /^\s*(?:(?:\(?[0-9A-Za-z]{1,3}\)?[.)\u3001\uff1f])|(?:[\uff08(][\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]{1,3}[\uff09)]))\s+(.+)$/,
  );
  if (orderedMatch?.[1]) {
    const contentStart = orderedMatch[0].length - orderedMatch[1].length;
    return {
      ordered: true,
      text: normalizeBlockText(orderedMatch[1]),
      inlines: sliceLineInlines(line, contentStart),
      style: line.style,
    };
  }

  const unorderedMatch = line.text.match(
    /^\s*(?:[\u2022\u25cf\u25cb\u25a0\u25c6]|[-*])\s+(.+)$/,
  );
  if (unorderedMatch?.[1]) {
    return {
      ordered: false,
      text: normalizeBlockText(unorderedMatch[1]),
      style: line.style,
    };
  }

  return undefined;
}

function inferParagraphStyle(
  role: DocParagraphBlock['role'],
  _text: string,
): DocTextStyle {
  if (role === 'title') {
    return {
      fontSize: 22,
      fontWeight: 700,
      lineHeight: 1.45,
      color: '#111827',
      textAlign: 'left',
      fontFamily: DOC_FONT_FAMILY,
      spacingBefore: 30,
      spacingAfter: 55,
      paddingBottom: 4,
    };
  }

  if (role === 'heading') {
    return {
      fontSize: 16,
      fontWeight: 700,
      lineHeight: 1.65,
      color: '#1f2937',
      textAlign: 'left',
      fontFamily: DOC_FONT_FAMILY,
      spacingAfter: 16,
      paddingTop: 4,
      paddingBottom: 4,
    };
  }

  return {
    fontSize: 14,
    fontWeight: 400,
    lineHeight: 1.8,
    color: '#111827',
    textAlign: 'left',
    fontFamily: DOC_FONT_FAMILY,
    spacingAfter: 18,
  };
}

function inferListStyle(ordered: boolean): DocTextStyle {
  return {
    fontSize: 14,
    fontWeight: 400,
    lineHeight: 1.7,
    color: '#111827',
    textAlign: 'left',
    fontFamily: DOC_FONT_FAMILY,
    paddingLeft: ordered ? 2 : 0,
  };
}

function inferTableStyle(): DocTableStyle {
  return {
    headerBackgroundColor: '#eef4ff',
    headerTextColor: '#1d4ed8',
    borderColor: '#cbd5e1',
    cellBackgroundColor: '#ffffff',
    stripedRowBackgroundColor: '#f8fafc',
  };
}

function estimateTableColumns(rows: PendingTableRow[]) {
  const columnCount = Math.max(...rows.map((row) => row.cells.length), 1);
  const weights = Array.from({ length: columnCount }, (_, columnIndex) =>
    Math.max(
      8,
      ...rows.map((row) => {
        const text = row.cells[columnIndex]?.text ?? '';
        return Array.from(text).reduce(
          (sum, char) => sum + (/[\u4e00-\u9fa5]/.test(char) ? 2 : 1),
          0,
        );
      }),
    ),
  );
  const total = weights.reduce((sum, weight) => sum + weight, 0) || 1;
  const availableWidth =
    DEFAULT_DOC_PAGE.width -
    DEFAULT_DOC_PAGE.marginLeft -
    DEFAULT_DOC_PAGE.marginRight;
  return weights.map((weight) =>
    Math.max(64, (weight / total) * availableWidth),
  );
}

function createParagraphBlock(
  text: string,
  index: number,
  inlines?: DocTextInline[],
  style?: DocTextStyle,
  pageBreakBefore?: boolean,
  outlineLevel?: number,
  isTableOfContents?: boolean,
  tocSpacingAfter?: number,
): DocParagraphBlock {
  const compactLength = text.replace(/\s+/g, '').length;
  const hasImages = Boolean(inlines?.some((inline) => inline.type === 'image'));
  // 已解析到段落样式时，仅在大纲级别或标题字体给出明确信号后才判为标题，
  // 避免把普通短正文误判为 heading 并丢弃行内混合字号。
  const hasHeadingStyleSignal =
    outlineLevel !== undefined ||
    (style?.fontWeight ?? 0) >= 600 ||
    (style?.fontSize ?? 0) >= 18;
  const role =
    index === 0 && compactLength <= 24
      ? 'title'
      : compactLength > 0 &&
        compactLength <= 18 &&
        !hasImages &&
        !isTableOfContents &&
        !/[|:\uff1a]/.test(text) &&
        !/[0-9]{4,}/.test(text) &&
        (!style || hasHeadingStyleSignal)
      ? 'heading'
      : 'body';
  const inferredStyle = isTableOfContents
    ? {
        ...inferParagraphStyle('body', text),
        fontSize: 14,
        fontWeight: 400,
        lineHeight: 1.5,
        // 短目录保留 31.2px 目录网格；仅在整段目录超出单页时收敛段后距。
        spacingAfter: tocSpacingAfter ?? 10.2,
      }
    : inferParagraphStyle(role, text);
  const mergedStyle = mergeStyleIntoTextStyle(inferredStyle, style);
  if (style) {
    // 已解析到 Word 段落属性时不叠加推断标题的 CSS 内边距，避免连续标题累计推迟分页。
    mergedStyle.paddingTop = style.paddingTop;
    mergedStyle.paddingBottom = style.paddingBottom;
  }
  if (
    style &&
    role === 'body' &&
    !isTableOfContents &&
    style.spacingAfter === undefined
  ) {
    // Word 与浏览器字体行框不同；保留适度排版补偿，同时避免套用标题级的大间距。
    mergedStyle.spacingAfter = 12;
  }
  const shouldUseSourceLineMultiplier =
    Boolean(isTableOfContents) ||
    ((mergedStyle.fontSize ?? 0) >= 28 && mergedStyle.textAlign === 'center');
  if (
    shouldUseSourceLineMultiplier &&
    mergedStyle.lineHeightMultiplier !== undefined
  ) {
    // 大字号居中标题与目录依赖 Word 的倍数行距；直接使用绝对行距会压住文字或撑大目录。
    mergedStyle.lineHeight = mergedStyle.lineHeightMultiplier;
  }

  return {
    id: `doc-p-${index + 1}`,
    type: 'paragraph',
    text,
    inlines,
    role,
    style: mergedStyle,
    pageBreakBefore,
    outlineLevel,
    isTableOfContents: isTableOfContents || undefined,
  };
}

/** 把列宽转换为从零开始的边界坐标，用于对齐不同物理网格的表格行。 */
function tableColumnBoundaries(columns: number[]) {
  return columns.reduce(
    (boundaries, width) => [
      ...boundaries,
      boundaries[boundaries.length - 1] + width,
    ],
    [0],
  );
}

/** 按 TC80 合并标志及行边界把物理网格转换为 HTML 单元格。 */
function normalizeTableRows(rows: PendingTableRow[], masterColumns: number[]) {
  const masterBoundaries = tableColumnBoundaries(masterColumns);
  return rows.map((row, rowIndex) => ({
    ...row,
    cells: row.cells.flatMap((cell, columnIndex) => {
      const layout = row.cellLayouts?.[columnIndex];
      if (
        layout?.horizontalMerge === 'continue' ||
        layout?.verticalMerge === 'continue'
      ) {
        return [];
      }

      let colSpan = 1;
      if (layout?.horizontalMerge === 'restart') {
        while (
          row.cellLayouts?.[columnIndex + colSpan]?.horizontalMerge ===
          'continue'
        ) {
          colSpan += 1;
        }
      }

      const rowBoundaries = row.columns
        ? tableColumnBoundaries(row.columns)
        : undefined;
      const sourceRight = rowBoundaries?.[columnIndex + colSpan];
      const masterLeftIndex = rowBoundaries
        ? masterBoundaries.findIndex(
            (boundary) => Math.abs(boundary - rowBoundaries[columnIndex]) < 1,
          )
        : -1;
      const masterRightIndex =
        sourceRight === undefined
          ? -1
          : masterBoundaries.findIndex(
              (boundary) => Math.abs(boundary - sourceRight) < 1,
            );
      const gridSpan =
        masterLeftIndex >= 0 && masterRightIndex > masterLeftIndex
          ? masterRightIndex - masterLeftIndex
          : colSpan;

      let rowSpan = 1;
      if (layout?.verticalMerge === 'restart') {
        while (
          rows[rowIndex + rowSpan]?.cellLayouts?.[columnIndex]
            ?.verticalMerge === 'continue'
        ) {
          rowSpan += 1;
        }
      }

      return [
        {
          ...cell,
          colSpan: gridSpan > 1 ? gridSpan : cell.colSpan,
          rowSpan: rowSpan > 1 ? rowSpan : cell.rowSpan,
          verticalAlign: layout?.verticalAlign,
        },
      ];
    }),
  }));
}

function createTableBlock(
  rows: PendingTableRow[],
  index: number,
  spacingBefore?: number,
  spacingAfter?: number,
  tableGridLineHeight?: number,
  documentGridLineHeight?: number,
): DocTableBlock {
  const structuralColumns = rows
    .map((row) => row.columns)
    .filter((columns): columns is number[] => Boolean(columns?.length))
    .sort(
      (left, right) =>
        right.length - left.length ||
        right.reduce((sum, width) => sum + width, 0) -
          left.reduce((sum, width) => sum + width, 0),
    )[0];
  // 二进制 DOC 已提供真实表格网格时不套用降级主题，避免虚构蓝色表头和斑马纹。
  const tableStyle: DocTableStyle = structuralColumns
    ? {
        headerBackgroundColor: '#ffffff',
        headerTextColor: '#111827',
        borderColor: '#000000',
        cellBackgroundColor: '#ffffff',
        stripedRowBackgroundColor: '#ffffff',
      }
    : inferTableStyle();
  const columns = structuralColumns ?? estimateTableColumns(rows);
  const verticalCellPadding = structuralColumns ? 3.5 : 5;
  const horizontalCellPadding = structuralColumns ? 6.25 : 8;
  const normalizedRows = normalizeTableRows(rows, columns);
  const width =
    rows.find((row) => row.width !== undefined)?.width ??
    (structuralColumns
      ? structuralColumns.reduce((sum, columnWidth) => sum + columnWidth, 0)
      : undefined);
  const firstRowHasMerges = Boolean(
    rows[0]?.cellLayouts?.some(
      (layout) => layout.horizontalMerge || layout.verticalMerge,
    ),
  );
  const hasHeaderRow =
    !structuralColumns &&
    !firstRowHasMerges &&
    normalizedRows[0]?.cells.length === columns.length &&
    normalizedRows[0].cells.length >= 2 &&
    normalizedRows[0].cells.every((cell) => Boolean(cell.text));
  return {
    id: `doc-table-${index + 1}`,
    type: 'table',
    style: tableStyle,
    columns,
    width,
    align: rows.find((row) => row.align)?.align,
    offsetLeft: rows.find((row) => row.offsetLeft !== undefined)?.offsetLeft,
    spacingBefore,
    spacingAfter,
    rows: normalizedRows.map((row, rowIndex) => ({
      id: `doc-table-${index + 1}-row-${rowIndex + 1}`,
      height: row.height,
      heightRule: row.heightRule,
      cells: row.cells.map((cell, cellIndex) => {
        const sourceStyle = tableCellTextStyle(cell.style);
        // 禁用文档网格的表格使用 Word 自动单倍行距；其余结构化表格跟随节内网格。
        const resolvedGridLineHeight =
          sourceStyle?.useDocumentGrid === false
            ? (sourceStyle.fontSize ?? 13) * 1.29
            : structuralColumns
            ? documentGridLineHeight ?? tableGridLineHeight
            : tableGridLineHeight ?? documentGridLineHeight;
        return {
          id: `doc-table-${index + 1}-cell-${rowIndex + 1}-${cellIndex + 1}`,
          text: cell.text,
          inlines: cell.inlines,
          colSpan: cell.colSpan,
          rowSpan: cell.rowSpan,
          verticalAlign: cell.verticalAlign,
          style: {
            color:
              rowIndex === 0 && hasHeaderRow
                ? tableStyle.headerTextColor
                : '#111827',
            backgroundColor:
              rowIndex === 0 && hasHeaderRow
                ? tableStyle.headerBackgroundColor
                : rowIndex % 2 === 1
                ? tableStyle.cellBackgroundColor
                : tableStyle.stripedRowBackgroundColor,
            fontSize: rowIndex === 0 ? 13 : 13,
            fontWeight: rowIndex === 0 && hasHeaderRow ? 700 : 400,
            lineHeight: resolvedGridLineHeight
              ? resolvedGridLineHeight
              : sourceStyle?.lineHeight ?? 1.25,
            fontFamily: DOC_FONT_FAMILY,
            paddingTop: resolvedGridLineHeight ? 0 : verticalCellPadding,
            paddingRight: horizontalCellPadding,
            paddingBottom: resolvedGridLineHeight ? 0 : verticalCellPadding,
            paddingLeft: horizontalCellPadding,
            ...sourceStyle,
            ...(resolvedGridLineHeight
              ? {
                  lineHeight: resolvedGridLineHeight,
                  paddingTop: 0,
                  paddingBottom: 0,
                }
              : undefined),
          },
        };
      }),
    })),
  };
}

function createListBlock(items: ParsedListLine[], index: number): DocListBlock {
  const orderedCount = items.filter((item) => item.ordered).length;
  const ordered = orderedCount >= items.length / 2;

  return {
    id: `doc-list-${index + 1}`,
    type: 'list',
    ordered,
    style: mergeStyleIntoTextStyle(
      inferListStyle(ordered),
      items.find((item) => item.style)?.style,
    ),
    items: items.map((item, itemIndex) => ({
      id: `doc-list-${index + 1}-item-${itemIndex + 1}`,
      text: item.text,
      inlines: item.inlines,
    })),
  };
}

function dominantStyle(segments: DocTextSegment[]) {
  return segments.reduce<DocTextStyle | undefined>((style, segment) => {
    // 字符高亮只能作用于对应 inline，不能扩散成整段背景。
    const {
      backgroundColor: _backgroundColor,
      paragraphBackgroundColor,
      ...paragraphStyle
    } = segment.style ?? {};
    return mergeTextStyle(style, {
      ...paragraphStyle,
      backgroundColor: paragraphBackgroundColor,
    });
  }, undefined);
}

async function blocksFromSegments(
  segments: DocTextSegment[],
  images: DocImage[] = [],
  options: DocBlockBuildOptions,
  drawingImages: Array<DocImage | undefined> = [],
): Promise<DocBlock[]> {
  const pendingTableRows: PendingTableRow[] = [];
  const pendingTableCells: PendingTableCell[] = [];
  const pendingListItems: ParsedListLine[] = [];
  const normalizedSegments = normalizeDocTextSegments(
    segments,
    images,
    drawingImages,
  );
  const tocParagraphSegments = normalizedSegments.filter(
    (segment) => segment.text === '\n' && segment.isTableOfContents,
  );
  const tocLineHeightTotal = tocParagraphSegments.reduce((total, segment) => {
    const fontSize = segment.style?.fontSize ?? 14;
    const lineHeight = segment.style?.lineHeight ?? 1.5;
    return total + (lineHeight > 4 ? lineHeight : fontSize * lineHeight);
  }, 0);
  const tocFixedAllowance =
    (options.defaultGridLineHeight ?? 31.2) * 2 +
    (options.documentGridLineHeight ?? 20.8) * 4;
  const tocSpacingAfter =
    options.pageContentHeight !== undefined &&
    tocParagraphSegments.length > 0 &&
    tocLineHeightTotal +
      tocParagraphSegments.length * 10.2 +
      tocFixedAllowance >
      options.pageContentHeight
      ? Math.max(
          0,
          (options.pageContentHeight - tocFixedAllowance - tocLineHeightTotal) /
            tocParagraphSegments.length,
        )
      : undefined;
  const builder = new DocBlockStreamBuilder({
    onBatch: options.onBatch
      ? ({ startIndex, blocks }) => options.onBatch!(startIndex, blocks)
      : undefined,
  });
  let currentLine = '';
  let currentLineInlines: DocTextInline[] = [];
  let currentLineSegments: DocTextSegment[] = [];
  let pendingTableColumns: number[] | undefined;
  let pendingTableAlign: DocTableBlock['align'];
  let pendingTableOffsetLeft: number | undefined;
  let pendingTableWidth: number | undefined;
  let pendingTableCellLayouts: DocTableCellLayout[] | undefined;
  let pendingTableRowHeight: number | undefined;
  let pendingTableRowHeightRule: 'atLeast' | 'exact' | undefined;
  let pendingTableSpacingBefore: number | undefined;
  let pendingBlockSpacingBefore = 0;

  /** 空段落只在紧邻图片时折算为段前距，避免全局保留空块扰乱估算分页。 */
  const emptyParagraphHeight = (style?: DocTextStyle) => {
    const fontSize = style?.fontSize ?? 14;
    const lineHeight = style?.lineHeight ?? 1.8;
    return lineHeight > 4 ? lineHeight : fontSize * lineHeight;
  };

  const makeLine = (boundary?: DocImageSegment): DocLine => {
    const text = currentLine;
    const structuralSegments = boundary
      ? [...currentLineSegments, boundary]
      : currentLineSegments;
    const inTable = structuralSegments.some((segment) => segment.inTable);
    const isTableOfContents = structuralSegments.some(
      (segment) => segment.isTableOfContents,
    );
    const sourceStyle = dominantStyle(currentLineSegments);
    const fontSize = sourceStyle?.fontSize ?? 14;
    const largestFontSize = Math.max(
      fontSize,
      ...currentLineSegments.map((segment) => segment.style?.fontSize ?? 0),
    );
    const explicitLineHeight =
      sourceStyle?.lineHeight === undefined
        ? undefined
        : sourceStyle.lineHeight > 4
        ? sourceStyle.lineHeight
        : fontSize * sourceStyle.lineHeight;
    const minimumLineHeight = Math.max(
      largestFontSize * 1.06,
      sourceStyle?.useDocumentGrid === false
        ? 0
        : options.defaultGridLineHeight ?? 0,
    );
    const documentGridLineHeight = options.documentGridLineHeight;
    const naturalWordLineHeight = largestFontSize * 1.3;
    const snappedGridLineHeight =
      sourceStyle?.useDocumentGrid !== false && documentGridLineHeight
        ? Math.ceil(
            Math.max(0, naturalWordLineHeight - 0.01) / documentGridLineHeight,
          ) * documentGridLineHeight
        : undefined;
    const rawGridPadding =
      snappedGridLineHeight !== undefined && documentGridLineHeight
        ? snappedGridLineHeight -
          naturalWordLineHeight -
          documentGridLineHeight / 2
        : 0;
    // Word 会把跨越多条文档网格的正文居中放进固定网格槽；浏览器只按当前最大字号生成行盒。
    // 将缺失的上下留白补回段落流，可避免删减大字号片段后把后续内容整体向上拉动。
    const gridPadding = rawGridPadding > 0.5 ? rawGridPadding + 0.75 : 0;
    const style =
      !inTable &&
      !isTableOfContents &&
      text.trim().length > 0 &&
      minimumLineHeight > 0 &&
      (explicitLineHeight === undefined ||
        explicitLineHeight < minimumLineHeight)
        ? {
            ...sourceStyle,
            lineHeight: minimumLineHeight,
            // 文档网格已经承担正文节奏，未显式声明的段后距不能再套用浏览器补偿。
            spacingAfter: sourceStyle?.spacingAfter ?? 0,
            paddingTop: (sourceStyle?.paddingTop ?? 0) + gridPadding / 2,
            paddingBottom: (sourceStyle?.paddingBottom ?? 0) + gridPadding / 2,
          }
        : sourceStyle;
    return {
      text,
      inlines: mergeAdjacentInlines(trimInlines(currentLineInlines)),
      style,
      inTable,
      tableRowEnd: structuralSegments.some((segment) => segment.tableRowEnd),
      tableRowHeight: structuralSegments.find(
        (segment) => segment.tableRowHeight !== undefined,
      )?.tableRowHeight,
      tableRowHeightRule: structuralSegments.find(
        (segment) => segment.tableRowHeightRule !== undefined,
      )?.tableRowHeightRule,
      pageBreakBefore: structuralSegments.some(
        (segment) => segment.pageBreakBefore,
      ),
      outlineLevel: structuralSegments.find(
        (segment) => segment.outlineLevel !== undefined,
      )?.outlineLevel,
      isTableOfContents,
      listId: structuralSegments.find((segment) => segment.listId !== undefined)
        ?.listId,
      listLevel: structuralSegments.find(
        (segment) => segment.listLevel !== undefined,
      )?.listLevel,
      tableColumns: structuralSegments.find(
        (segment) => segment.tableColumns?.length,
      )?.tableColumns,
      tableAlign: structuralSegments.find(
        (segment) => segment.tableAlign !== undefined,
      )?.tableAlign,
      tableOffsetLeft: structuralSegments.find(
        (segment) => segment.tableOffsetLeft !== undefined,
      )?.tableOffsetLeft,
      tableWidth: structuralSegments.find(
        (segment) => segment.tableWidth !== undefined,
      )?.tableWidth,
      tableCellLayouts: structuralSegments.find(
        (segment) => segment.tableCellLayouts?.length,
      )?.tableCellLayouts,
      match: (pattern) => text.match(pattern),
    };
  };

  const resetLine = () => {
    currentLine = '';
    currentLineInlines = [];
    currentLineSegments = [];
  };

  const captureTableStructure = (line: DocLine) => {
    pendingTableColumns = line.tableColumns ?? pendingTableColumns;
    pendingTableAlign = line.tableAlign ?? pendingTableAlign;
    pendingTableOffsetLeft = line.tableOffsetLeft ?? pendingTableOffsetLeft;
    pendingTableWidth = line.tableWidth ?? pendingTableWidth;
    pendingTableCellLayouts = line.tableCellLayouts ?? pendingTableCellLayouts;
    pendingTableRowHeight = line.tableRowHeight ?? pendingTableRowHeight;
    pendingTableRowHeightRule =
      line.tableRowHeightRule ?? pendingTableRowHeightRule;
  };

  const commitTableRow = () => {
    if (!pendingTableCells.length) return;
    pendingTableRows.push({
      cells: [...pendingTableCells],
      columns: pendingTableColumns,
      align: pendingTableAlign,
      offsetLeft: pendingTableOffsetLeft,
      width: pendingTableWidth,
      cellLayouts: pendingTableCellLayouts,
      height: pendingTableRowHeight,
      heightRule: pendingTableRowHeightRule,
    });
    pendingTableCells.length = 0;
    pendingTableColumns = undefined;
    pendingTableAlign = undefined;
    pendingTableOffsetLeft = undefined;
    pendingTableWidth = undefined;
    pendingTableCellLayouts = undefined;
    pendingTableRowHeight = undefined;
    pendingTableRowHeightRule = undefined;
  };

  const flushTable = async (spacingAfter?: number) => {
    commitTableRow();
    if (!pendingTableRows.length) return;
    const rows = [...pendingTableRows];
    const spacingBefore = pendingTableSpacingBefore;
    pendingTableRows.length = 0;
    pendingTableSpacingBefore = undefined;
    if (rows.length === 1) {
      const text = rows[0].cells.map((cell) => cell.text).join(' ');
      const inlines = rows[0].cells.flatMap((cell) => cell.inlines);
      await builder.add(
        createParagraphBlock(
          text,
          builder.nextSourceIndex,
          inlines,
          spacingBefore ? { spacingBefore } : undefined,
        ),
      );
    } else {
      await builder.add(
        createTableBlock(
          rows,
          builder.nextSourceIndex,
          spacingBefore,
          spacingAfter,
          options.defaultGridLineHeight,
          options.documentGridLineHeight,
        ),
      );
    }
  };

  const flushList = async () => {
    if (!pendingListItems.length) return;
    const items = [...pendingListItems];
    pendingListItems.length = 0;
    if (items.length === 1) {
      await builder.add(
        createParagraphBlock(
          items[0].text,
          builder.nextSourceIndex,
          items[0].inlines,
          items[0].style,
        ),
      );
    } else {
      await builder.add(createListBlock(items, builder.nextSourceIndex));
    }
  };

  /** 二进制正文不存储自动编号文字，按 PlfLst/LFO 状态补回可见前缀。 */
  const applyAutomaticNumbering = (line: DocLine): DocLine => {
    if (
      line.inTable ||
      line.isTableOfContents ||
      line.listId === undefined ||
      line.listLevel === undefined ||
      !options.numbering
    ) {
      return line;
    }
    const prefix = nextDocNumberPrefix(
      options.numbering,
      line.listId,
      line.listLevel,
    );
    if (!prefix?.text) return line;
    const normalized = normalizeBlockText(line.text);
    if (
      normalized === prefix.text ||
      normalized.startsWith(`${prefix.text} `) ||
      normalized.startsWith(`${prefix.text}\t`)
    ) {
      return line;
    }
    const separator =
      prefix.suffix === 'space' ? ' ' : prefix.suffix === 'tab' ? '\t' : '';
    const prefixText = `${prefix.text}${separator}`;
    return {
      ...line,
      text: `${prefixText}${line.text}`,
      inlines: [
        { type: 'text', text: prefixText, style: line.style },
        ...line.inlines,
      ],
    };
  };

  const processLine = async (inputLine: DocLine) => {
    let line = applyAutomaticNumbering(inputLine);
    const textLine = normalizeBlockText(line.text);
    if (!textLine) {
      if (line.inTable) {
        if (!pendingTableRows.length && !pendingTableCells.length) {
          pendingTableSpacingBefore = pendingBlockSpacingBefore || undefined;
          pendingBlockSpacingBefore = 0;
        }
        captureTableStructure(line);
        pendingTableCells.push(...splitTableCells(line));
        if (line.tableRowEnd && pendingTableCells.length) {
          commitTableRow();
        }
        return;
      }
      const followsTable =
        pendingTableRows.length > 0 || pendingTableCells.length > 0;
      // 表格后的空段落归入表格尾距；其余空段落累计到下一个可见内容块。
      await flushTable(
        followsTable
          ? options.defaultGridLineHeight ?? emptyParagraphHeight(line.style)
          : undefined,
      );
      await flushList();
      if (line.inlines.some((inline) => inline.type === 'image')) {
        const imageStyle =
          pendingBlockSpacingBefore > 0
            ? {
                ...line.style,
                spacingBefore:
                  (line.style?.spacingBefore ?? 0) + pendingBlockSpacingBefore,
              }
            : line.style;
        await builder.add(
          createParagraphBlock(
            '',
            builder.nextSourceIndex,
            line.inlines,
            imageStyle,
          ),
        );
        pendingBlockSpacingBefore = 0;
      } else if (!followsTable) {
        pendingBlockSpacingBefore +=
          options.defaultGridLineHeight ?? emptyParagraphHeight(line.style);
      }
      return;
    }

    if (line.inTable) {
      await flushList();
      if (!pendingTableRows.length && !pendingTableCells.length) {
        pendingTableSpacingBefore = pendingBlockSpacingBefore || undefined;
        pendingBlockSpacingBefore = 0;
      }
      captureTableStructure(line);
      pendingTableCells.push(...splitTableCells(line));
      if (line.tableRowEnd && pendingTableCells.length) {
        commitTableRow();
      }
      return;
    }

    if (looksLikeTableRow(textLine)) {
      await flushList();
      if (!pendingTableRows.length) {
        pendingTableSpacingBefore = pendingBlockSpacingBefore || undefined;
        pendingBlockSpacingBefore = 0;
      }
      // 旧 DOC/WPS 即使缺失行结束标志，行尾 PAPX 仍可能携带真实列宽和对齐信息。
      pendingTableRows.push({
        cells: splitTableCells(line),
        columns: line.tableColumns,
        align: line.tableAlign,
        offsetLeft: line.tableOffsetLeft,
        width: line.tableWidth,
        cellLayouts: line.tableCellLayouts,
        height: line.tableRowHeight,
        heightRule: line.tableRowHeightRule,
      });
      return;
    }

    if (pendingBlockSpacingBefore > 0) {
      line = {
        ...line,
        style: {
          ...line.style,
          spacingBefore:
            (line.style?.spacingBefore ?? 0) + pendingBlockSpacingBefore,
        },
      };
      pendingBlockSpacingBefore = 0;
    }

    if (line.outlineLevel !== undefined) {
      // 源大纲语义优先于列表外观推断，避免编号标题被降成普通列表项。
      await flushTable();
      await flushList();
      await builder.add(
        createParagraphBlock(
          textLine,
          builder.nextSourceIndex,
          line.inlines,
          line.style,
          line.pageBreakBefore,
          line.outlineLevel,
          line.isTableOfContents,
          line.isTableOfContents ? tocSpacingAfter : undefined,
        ),
      );
      return;
    }

    const listLine = parseListLine(line);
    if (listLine) {
      await flushTable();
      if (!listLine.inlines?.length) {
        listLine.inlines = line.inlines;
      }
      pendingListItems.push(listLine);
      return;
    }

    await flushTable();
    await flushList();
    await builder.add(
      createParagraphBlock(
        textLine,
        builder.nextSourceIndex,
        line.inlines,
        line.style,
        line.pageBreakBefore,
        line.outlineLevel,
        line.isTableOfContents,
        line.isTableOfContents ? tocSpacingAfter : undefined,
      ),
    );
  };

  for (let index = 0; index < normalizedSegments.length; index += 1) {
    const segment = normalizedSegments[index];
    if (segment.text === '\f') {
      const line = makeLine(segment);
      resetLine();
      await processLine(line);
      await flushTable();
      await flushList();
      pendingBlockSpacingBefore = 0;
      // 旧版 DOC 的 0x0C 是强制分页，使用隐藏占位块把分页语义传给渲染器。
      await builder.add({
        ...createParagraphBlock('', builder.nextSourceIndex),
        pageBreakBefore: true,
      });
    } else if (segment.text === '\n') {
      const insideTable =
        segment.inTable ||
        currentLineSegments.some((lineSegment) => lineSegment.inTable);
      if (insideTable) {
        // 表格单元格内的段落标记属于单元格换行，真正的行结束由 fTtp 标志决定。
        currentLine += '\n';
        currentLineInlines.push({
          type: 'text',
          text: '\n',
          style: segment.style,
        });
        currentLineSegments.push(segment);
        continue;
      }
      const line = makeLine(segment);
      resetLine();
      await processLine(line);
    } else if (segment.tableRowEnd) {
      const line = makeLine(segment);
      resetLine();
      await processLine(line);
    } else if (segment.image) {
      currentLineInlines.push({ type: 'image', image: segment.image });
      // 图片锚点同样携带所在段落的对齐与间距，不能只保留图片资源本身。
      currentLineSegments.push(segment);
    } else if (
      segment.text === '|' &&
      currentLine.endsWith('|') &&
      !segment.inTable &&
      !currentLineSegments.some((lineSegment) => lineSegment.inTable)
    ) {
      currentLine = currentLine.slice(0, -1);
      const previousInline = currentLineInlines[currentLineInlines.length - 1];
      if (previousInline?.type === 'text') {
        previousInline.text = previousInline.text.slice(0, -1);
        if (!previousInline.text) currentLineInlines.pop();
      }
      // 第二个单元格结束符仍携带行级 PAPX，作为边界传入才能保住真实表格网格。
      const line = makeLine(segment);
      resetLine();
      await processLine(line);
    } else {
      currentLine += segment.text;
      currentLineInlines.push({
        type: 'text',
        text: segment.text,
        style: segment.style,
      });
      currentLineSegments.push(segment);
    }

    if ((index + 1) % 64 === 0) {
      await options.checkpoint({
        stage: 'content',
        completed: index + 1,
        total: normalizedSegments.length,
        percent:
          0.7 + ((index + 1) / Math.max(1, normalizedSegments.length)) * 0.22,
        message: '正在解析 DOC 正文内容',
      });
    }
  }

  await processLine(makeLine());
  await flushTable();
  await flushList();
  return builder.finish();
}

async function blocksFromText(
  text: string,
  options: DocBlockBuildOptions,
): Promise<DocBlock[]> {
  return blocksFromSegments([{ text }], [], options);
}

/** 提取并汇总 `extractImageAt` 返回的数据。 */
function extractImageAt(bytes: Uint8Array, start: number) {
  if (
    bytes[start] === 0x89 &&
    bytes[start + 1] === 0x50 &&
    bytes[start + 2] === 0x4e &&
    bytes[start + 3] === 0x47 &&
    bytes[start + 4] === 0x0d &&
    bytes[start + 5] === 0x0a &&
    bytes[start + 6] === 0x1a &&
    bytes[start + 7] === 0x0a
  ) {
    let offset = start + 8;
    while (offset + 12 <= bytes.length) {
      const view = new DataView(
        bytes.buffer,
        bytes.byteOffset,
        bytes.byteLength,
      );
      const chunkLength = readUint32BE(view, offset);
      const chunkType = String.fromCharCode(
        bytes[offset + 4],
        bytes[offset + 5],
        bytes[offset + 6],
        bytes[offset + 7],
      );
      const nextOffset = offset + 12 + chunkLength;
      if (nextOffset > bytes.length) break;
      offset = nextOffset;
      if (chunkType === 'IEND') {
        return {
          mimeType: 'image/png',
          bytes: bytes.slice(start, offset),
        };
      }
    }
  }

  if (
    bytes[start] === 0xff &&
    bytes[start + 1] === 0xd8 &&
    bytes[start + 2] === 0xff
  ) {
    for (let index = start + 2; index + 1 < bytes.length; index += 1) {
      if (bytes[index] === 0xff && bytes[index + 1] === 0xd9) {
        return {
          mimeType: 'image/jpeg',
          bytes: bytes.slice(start, index + 2),
        };
      }
    }
  }

  return undefined;
}

function readImageSize(bytes: Uint8Array, mimeType: string) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  if (mimeType === 'image/png' && bytes.length >= 24) {
    return {
      width: readUint32BE(view, 16),
      height: readUint32BE(view, 20),
    };
  }

  if (mimeType === 'image/jpeg') {
    let offset = 2;
    while (offset + 9 < bytes.length) {
      if (bytes[offset] !== 0xff) {
        offset += 1;
        continue;
      }
      const marker = bytes[offset + 1];
      const length = readUint16BE(view, offset + 2);
      if (marker >= 0xc0 && marker <= 0xc3 && offset + 8 < bytes.length) {
        return {
          width: readUint16BE(view, offset + 7),
          height: readUint16BE(view, offset + 5),
        };
      }
      offset += Math.max(2, length + 2);
    }
  }

  return {};
}

/** 从 DOC 的 PICF 记录恢复随文图片经过缩放后的最终显示尺寸。 */
function readInlinePictureLayouts(bytes: Uint8Array) {
  const layouts: Array<{
    start: number;
    end: number;
    width: number;
    height: number;
  }> = [];
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  for (let offset = 0; offset + 68 <= bytes.length; offset += 1) {
    const recordLength = view.getInt32(offset, true);
    const headerLength = readUint16(view, offset + 4);
    const mappingMode = readInt16(view, offset + 6);
    if (
      recordLength < 68 ||
      offset + recordLength > bytes.length ||
      headerLength !== 0x44 ||
      (mappingMode !== 0x64 && mappingMode !== 0x66) ||
      readUint16(view, offset + 66) !== 0
    ) {
      continue;
    }

    const initialWidth = readInt16(view, offset + 28);
    const initialHeight = readInt16(view, offset + 30);
    const horizontalScale = readUint16(view, offset + 32);
    const verticalScale = readUint16(view, offset + 34);
    if (
      initialWidth <= 0 ||
      initialHeight <= 0 ||
      horizontalScale <= 0 ||
      verticalScale <= 0
    ) {
      continue;
    }

    layouts.push({
      start: offset,
      end: offset + recordLength,
      // PICMID 的缩放值以千分之一表示，目标尺寸以 twip 表示。
      width: twipToPx((initialWidth * horizontalScale) / 1000),
      height: twipToPx((initialHeight * verticalScale) / 1000),
    });
    offset += recordLength - 1;
  }

  return layouts;
}

function textNearBytes(
  bytes: Uint8Array,
  start: number,
  before = 320,
  after = 80,
) {
  const slice = bytes.slice(
    Math.max(0, start - before),
    Math.min(bytes.length, start + after),
  );
  return Array.from(slice, (value) =>
    value >= 32 && value <= 126 ? String.fromCharCode(value) : ' ',
  ).join('');
}

function isLikelySameImageObject(
  left: DocImageCandidate,
  right: DocImageCandidate,
) {
  if (left.streamName !== right.streamName) return false;
  if (
    left.mimeType !== right.mimeType ||
    !left.width ||
    !left.height ||
    !right.width ||
    !right.height
  )
    return false;
  if (left.width !== right.width || left.height !== right.height) return false;

  const byteDelta = Math.abs(left.byteLength - right.byteLength);
  const isCloseLength =
    byteDelta <= 1024 ||
    byteDelta / Math.max(left.byteLength, right.byteLength) <= 0.02;
  const isOfficePreviewPair =
    (left.packagedMedia && right.webExtensionPreview) ||
    (left.webExtensionPreview && right.packagedMedia);
  const isNearAlternatePreview = Math.abs(left.offset - right.offset) <= 120000;

  return isCloseLength && isOfficePreviewPair && isNearAlternatePreview;
}

function chooseBetterImageCandidate(
  left: DocImageCandidate,
  right: DocImageCandidate,
) {
  if (left.packagedMedia !== right.packagedMedia)
    return left.packagedMedia ? left : right;
  if (left.byteLength !== right.byteLength)
    return left.byteLength > right.byteLength ? left : right;
  return left.offset <= right.offset ? left : right;
}

/** 将输入标准化为 `normalizeImageCandidates` 返回的结构。 */
function normalizeImageCandidates(
  candidates: DocImageCandidate[],
  resources: PortableResource[],
) {
  const normalized: DocImageCandidate[] = [];

  candidates.forEach((candidate) => {
    const duplicateIndex = normalized.findIndex((image) =>
      isLikelySameImageObject(image, candidate),
    );
    if (duplicateIndex === -1) {
      normalized.push(candidate);
      return;
    }

    normalized[duplicateIndex] = chooseBetterImageCandidate(
      normalized[duplicateIndex],
      candidate,
    );
  });

  return normalized
    .sort((left, right) => left.offset - right.offset)
    .map(
      (
        {
          byteLength,
          packagedMedia,
          webExtensionPreview,
          streamName,
          bytes,
          ...image
        },
        index,
      ) => {
        const resourceId = `doc:image:${index + 1}`;
        const resourceBytes = bytes.slice();
        resources.push({
          id: resourceId,
          encoding: 'binary',
          mimeType: image.mimeType,
          buffer: resourceBytes.buffer,
        });
        return {
          ...image,
          id: `doc-image-${index + 1}`,
          src: createResourceReference(resourceId),
        };
      },
    );
}

/** 提取并汇总 `extractDocImagesFromStream` 返回的数据。 */
function extractDocImagesFromStream(bytes: Uint8Array, streamName: string) {
  const candidates: DocImageCandidate[] = [];
  const seen = new Set<string>();
  const pictureLayouts = readInlinePictureLayouts(bytes);
  let pictureLayoutIndex = 0;
  const signatures = [
    { mimeType: 'image/png', header: [0x89, 0x50, 0x4e, 0x47] },
    { mimeType: 'image/jpeg', header: [0xff, 0xd8, 0xff] },
  ];

  for (let index = 0; index < bytes.length - 4; index += 1) {
    const signature = signatures.find(({ header }) =>
      header.every(
        (value, headerIndex) => bytes[index + headerIndex] === value,
      ),
    );
    if (!signature) continue;

    const extracted = extractImageAt(bytes, index);
    if (!extracted || extracted.bytes.length < 128) continue;

    const head = Array.from(extracted.bytes.slice(0, 16)).join(',');
    const tail = Array.from(
      extracted.bytes.slice(Math.max(0, extracted.bytes.length - 16)),
    ).join(',');
    const key = `${extracted.mimeType}:${extracted.bytes.length}:${head}:${tail}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const context = textNearBytes(bytes, index);
    while (
      pictureLayoutIndex < pictureLayouts.length &&
      pictureLayouts[pictureLayoutIndex].end <= index
    ) {
      pictureLayoutIndex += 1;
    }
    const pictureLayout = pictureLayouts[pictureLayoutIndex];
    const displaySize =
      pictureLayout &&
      pictureLayout.start <= index &&
      pictureLayout.end >= index + extracted.bytes.length
        ? {
            width: pictureLayout.width,
            height: pictureLayout.height,
          }
        : readImageSize(extracted.bytes, extracted.mimeType);

    candidates.push({
      bytes: extracted.bytes,
      mimeType: extracted.mimeType,
      offset: index,
      byteLength: extracted.bytes.length,
      packagedMedia: /drs\/media|drs\\media/.test(context),
      webExtensionPreview: /drs\/webExtensions|drs\\webExtensions/.test(
        context,
      ),
      streamName,
      ...displaySize,
    });
  }

  return candidates;
}

/** 提取并汇总 `extractDocImages` 返回的数据。 */
function extractDocImages(
  streams: Iterable<readonly [string, Uint8Array]>,
  resources: PortableResource[],
) {
  const candidates = Array.from(streams).flatMap(([streamName, stream]) =>
    extractDocImagesFromStream(stream, streamName),
  );

  return normalizeImageCandidates(candidates, resources);
}

/** 判断输入是否为随机 CFB Reader 提供的 DOC 核心流。 */
function isDocCoreStreamsInput(
  input: ArrayBuffer | Uint8Array | DocCoreStreamsInput,
): input is DocCoreStreamsInput {
  return (
    typeof input === 'object' &&
    input !== null &&
    'wordDocument' in input &&
    'tableStream' in input
  );
}

async function parsePlainLikeDoc(
  bytes: Uint8Array,
  fileName: string,
  warnings: string[],
  options: DocBlockBuildOptions,
  output?: DocCoreOutput,
): Promise<DocDocument> {
  const fullText = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
  const isRtf = fullText.trimStart().startsWith('{\\rtf');
  const text = isRtf
    ? fullText
        .replace(/\\'[0-9a-f]{2}/gi, '')
        .replace(/\\[a-z]+-?\d* ?/gi, '')
        .replace(/[{}]/g, '')
    : fullText.replace(/<[^>]+>/g, ' ');

  warnings.push(
    isRtf
      ? '\u68c0\u6d4b\u5230 RTF \u5185\u5bb9\uff0c\u5df2\u6309\u7eaf\u6587\u672c\u964d\u7ea7\u9884\u89c8\u3002'
      : '\u68c0\u6d4b\u5230\u975e OLE DOC \u5185\u5bb9\uff0c\u5df2\u6309\u7eaf\u6587\u672c\u964d\u7ea7\u9884\u89c8\u3002',
  );
  const metadataDocument = buildDocDocument(fileName, [], [...warnings]);
  await output?.documentMetadata(documentMetadataFromDoc(metadataDocument));
  const blocks = await blocksFromText(text, options);
  return buildDocDocument(fileName, blocks, warnings);
}

function buildDocDocument(
  fileName: string,
  blocks: DocBlock[],
  warnings: string[],
): DocDocument {
  const paragraphs = paragraphsFromDocBlocks(blocks);
  const title =
    paragraphs.find((paragraph) => paragraph.text)?.text ??
    (fileName || 'DOC \u6587\u6863');
  const images = [] as DocImage[];
  const outline = blocks.flatMap((block) =>
    block.type === 'paragraph' && block.outlineLevel !== undefined && block.text
      ? [
          {
            id: `outline-${block.id}`,
            text: block.text,
            level: block.outlineLevel,
            targetBlockId: block.id,
          },
        ]
      : [],
  );

  return {
    title,
    page: DEFAULT_DOC_PAGE,
    blocks,
    paragraphs,
    images,
    outline,
    warnings,
  };
}

/** 消费已提取图片，确保引用它们的 DOC 元数据和正文不会先到达。 */
async function flushDocResources(
  resources: PortableResource[],
  output: DocCoreOutput | undefined,
) {
  if (!output || !resources.length) return;
  const pending = resources.splice(0);
  for (const resource of pending) {
    await output.resource(resource);
  }
}

/** 解析 DOC/WPS 二进制，并返回环境无关的文档与图片资源。 */
export async function parseDocCore(
  input: ArrayBuffer | Uint8Array | DocCoreStreamsInput,
  context: DocCoreContext,
): Promise<DocCoreResult> {
  // 非 OLE 文件按纯文本降级处理；OLE DOC 则解析 CFB、FIB、piece table 和样式 run。
  const streamsInput = isDocCoreStreamsInput(input) ? input : undefined;
  const bytes = streamsInput
    ? undefined
    : input instanceof Uint8Array
    ? input
    : new Uint8Array(input as ArrayBuffer);
  const warnings: string[] = [];
  const resources: PortableResource[] = [];

  if (bytes && !isOleDoc(bytes)) {
    await context.checkpoint({
      stage: 'content',
      percent: 0.8,
      message: '正在解析 DOC 纯文本内容',
    });
    const document = await parsePlainLikeDoc(
      bytes,
      context.fileName,
      warnings,
      {
        checkpoint: context.checkpoint,
        onBatch: context.output
          ? (startIndex, blocks) =>
              context.output!.documentBlocks(startIndex, blocks)
          : undefined,
      },
      context.output,
    );
    await context.output?.documentMetadata(documentMetadataFromDoc(document));
    return { document, resources };
  }

  await context.checkpoint({
    stage: 'container',
    percent: 0.05,
    message: '正在读取 DOC 复合文档',
  });
  const cfb = bytes
    ? await parseCfb(bytes, {
        yieldIfNeeded: () => context.checkpoint(),
        allowPartialFinalSector: true,
      })
    : undefined;
  const inputStreams = streamsInput
    ? [...streamsInput.imageStreams]
    : undefined;
  const dataStream =
    inputStreams?.find(
      ([streamName]) =>
        streamName.replace(/^.*[\\/]/, '').toLowerCase() === 'data',
    )?.[1] ?? cfb?.getStream('Data');
  const wordDocument =
    streamsInput?.wordDocument ?? cfb?.getStream('WordDocument');

  if (!wordDocument) {
    throw new Error(
      'DOC \u6587\u4ef6\u7f3a\u5c11 WordDocument \u6570\u636e\u6d41',
    );
  }

  const fib = parseFib(wordDocument);
  const tableStream =
    streamsInput?.tableStream ?? cfb?.getStream(fib.tableStreamName);

  if (!tableStream) {
    throw new Error(
      `DOC \u6587\u4ef6\u7f3a\u5c11 ${fib.tableStreamName} \u6570\u636e\u6d41`,
    );
  }

  await context.checkpoint({
    stage: 'structure',
    percent: 0.25,
    message: '正在解析 DOC 文档结构',
  });
  const pieces = parsePieces(tableStream, fib);
  if (!pieces.length) {
    throw new Error(
      '\u6682\u672a\u80fd\u8bc6\u522b\u8be5 DOC \u6587\u4ef6\u7684\u6b63\u6587\u7247\u6bb5\u8868',
    );
  }

  const fonts = parseFontTable(tableStream, fib);
  const outlineCatalog = parseDocStyleOutlineCatalog(
    tableStream,
    fib.fcStshf,
    fib.lcbStshf,
  );
  const numbering = readDocNumberingCatalog(
    tableStream,
    fib.fcPlfLst,
    fib.lcbPlfLst,
    fib.fcPlfLfo,
    fib.lcbPlfLfo,
  );
  const sections = readDocSections(wordDocument, tableStream, fib);
  const dominantSection = [...sections].sort(
    (left, right) =>
      right.charEnd - right.charStart - (left.charEnd - left.charStart),
  )[0];
  const documentPage = dominantSection?.page ?? DEFAULT_DOC_PAGE;
  const documentGridLinePitch = sections.find(
    (section) => section.gridLinePitch !== undefined,
  )?.gridLinePitch;
  const normalStyle = readInheritedParagraphStyle(
    new Uint8Array([0, 0]),
    outlineCatalog,
    fonts,
  );
  const characterRuns = parseCharacterRuns(
    wordDocument,
    tableStream,
    fib,
    fonts,
  );
  const paragraphRuns = parseParagraphRuns(
    wordDocument,
    tableStream,
    fib,
    outlineCatalog,
    fonts,
    dataStream,
  );
  const paragraphLineMultiplierCounts = paragraphRuns.reduce((counts, run) => {
    const multiplier = run.style?.lineHeightMultiplier;
    if (multiplier !== undefined && !run.inTable && !run.isTableOfContents) {
      counts.set(multiplier, (counts.get(multiplier) ?? 0) + 1);
    }
    return counts;
  }, new Map<number, number>());
  const dominantParagraphLineMultiplier = [
    ...paragraphLineMultiplierCounts.entries(),
  ].sort((left, right) => right[1] - left[1])[0]?.[0];
  const defaultLineMultiplier =
    normalStyle?.lineHeightMultiplier ??
    (normalStyle?.lineHeight !== undefined && normalStyle.lineHeight <= 4
      ? normalStyle.lineHeight
      : dominantParagraphLineMultiplier);
  const defaultGridLineHeight =
    documentGridLinePitch !== undefined && defaultLineMultiplier !== undefined
      ? documentGridLinePitch * defaultLineMultiplier
      : undefined;
  warnings.push(...outlineCatalog.warnings);
  await context.checkpoint({
    stage: 'resources',
    percent: 0.5,
    message: '正在解析 DOC 图片资源',
  });
  const images = extractDocImages(inputStreams ?? cfb!.streams, resources);
  const headerStart = fib.ccpText + fib.ccpFtn;
  const headerPieces = slicePiecesByCharacterRange(
    pieces,
    headerStart,
    headerStart + fib.ccpHdr,
  );
  const headerText = textSegmentsFromPieces(
    wordDocument,
    headerPieces,
    characterRuns,
    paragraphRuns,
  )
    .map((segment) => segment.text)
    .join('');
  // 图片流按 Word story 顺序排列；页眉存在图片锚点时，第一张图不应再分配给正文。
  const headerImage = headerText.includes('\u0001') ? images[0] : undefined;
  const bodyImages = headerImage ? images.slice(1) : images;
  const footerPageNumbers = /\u0013PAGE\b/.test(headerText);
  const textBoxStart =
    fib.ccpText +
    fib.ccpFtn +
    fib.ccpHdr +
    fib.ccpMcr +
    fib.ccpAtn +
    fib.ccpEdn;
  const drawingTextBoxRanges = readDrawingTextBoxRanges(tableStream, fib);
  let textBoxes: DocDrawingTextBox[];
  if (drawingTextBoxRanges.length) {
    textBoxes = [];
    for (const range of drawingTextBoxRanges) {
      if (!range) {
        // PlcftxbxTxt 的空记录仍占用 ClientTextbox 索引，必须保留槽位。
        textBoxes.push({ text: '' });
        continue;
      }
      const scopedPieces = slicePiecesByCharacterRange(
        pieces,
        textBoxStart + range.charStart,
        textBoxStart + range.charEnd,
      );
      const scopedSegments = textSegmentsFromPieces(
        wordDocument,
        scopedPieces,
        characterRuns,
        paragraphRuns,
      );
      const scopedBlocks = await blocksFromSegments(scopedSegments, [], {
        checkpoint: context.checkpoint,
      });
      const paragraphs = scopedBlocks.filter(
        (block): block is DocParagraphBlock => block.type === 'paragraph',
      );
      const styledParagraph =
        paragraphs.find((paragraph) => paragraph.text.trim()) ?? paragraphs[0];
      textBoxes.push({
        text: paragraphs
          .map((paragraph) => paragraph.text.trim())
          .filter(Boolean)
          .join('\n'),
        style: styledParagraph?.style,
      });
    }
  } else {
    const textBoxPieces = slicePiecesByCharacterRange(
      pieces,
      textBoxStart,
      textBoxStart + fib.ccpTxbx,
    );
    const textBoxSegments = textSegmentsFromPieces(
      wordDocument,
      textBoxPieces,
      characterRuns,
      paragraphRuns,
    );
    const textBoxBlocks = await blocksFromSegments(textBoxSegments, [], {
      checkpoint: context.checkpoint,
    });
    textBoxes = textBoxBlocks.filter(
      (block): block is DocParagraphBlock => block.type === 'paragraph',
    );
  }
  const drawingCanvases = extractDocDrawingCanvases(
    tableStream,
    fib,
    textBoxes,
    resources,
    { sections, displayPage: documentPage },
  );
  const drawingImages = drawingCanvases.images;
  await flushDocResources(resources, context.output);
  const metadataDocument = buildDocDocument(
    context.fileName,
    [],
    [...warnings],
  );
  metadataDocument.page = documentPage;
  metadataDocument.images = [...drawingImages, ...images];
  metadataDocument.headerImage = headerImage;
  metadataDocument.footerPageNumbers = footerPageNumbers;
  await context.output?.documentMetadata(
    documentMetadataFromDoc(metadataDocument),
  );
  await context.checkpoint({
    stage: 'content',
    percent: 0.7,
    message: '正在解析 DOC 正文内容',
  });
  const mainPieces = slicePiecesByCharacterRange(pieces, 0, fib.ccpText);
  const segments = textSegmentsFromPieces(
    wordDocument,
    mainPieces,
    characterRuns,
    paragraphRuns,
  );
  const hasStructuralTableRows = paragraphRuns.some((run) => run.tableRowEnd);
  // 只有文档实际提供行结束标志时才采用 PAPX 表格结构；旧 WPS/DOC 常只有 inTable，
  // 此时继续使用单元格分隔符回退，避免把表格及其后的正文吞成同一行。
  const contentSegments = hasStructuralTableRows
    ? segments
    : segments.map((segment) => ({
        ...segment,
        inTable: undefined,
        tableRowEnd: undefined,
      }));
  const blocks = await blocksFromSegments(
    contentSegments,
    bodyImages,
    {
      checkpoint: context.checkpoint,
      numbering,
      defaultGridLineHeight,
      documentGridLineHeight: documentGridLinePitch,
      pageContentHeight:
        documentPage.minHeight -
        documentPage.marginTop -
        documentPage.marginBottom,
      onBatch: context.output
        ? (startIndex, batch) =>
            context.output!.documentBlocks(startIndex, batch)
        : undefined,
    },
    drawingCanvases.slots,
  );

  if (!blocks.length) {
    throw new Error(
      '\u8be5 DOC \u6587\u4ef6\u672a\u89e3\u6790\u5230\u53ef\u9884\u89c8\u6b63\u6587',
    );
  }

  warnings.push(
    drawingImages.length
      ? '已恢复 DOC/WPS 主文档中的 OfficeArt 绘图画布；分页仍由前端按源页面尺寸估算。'
      : images.length
      ? '当前为纯前端 DOC/WPS 降级预览，已提取到文档内图片，并按前端估算分页；暂未恢复精确锚点和复杂样式。'
      : '当前为纯前端 DOC/WPS 降级预览，已按前端估算分页；暂不还原复杂样式和图片锚点。',
  );
  await context.checkpoint({
    stage: 'assembling',
    percent: 0.95,
    message: '正在组装 DOC 文档',
  });
  const document = buildDocDocument(context.fileName, blocks, warnings);
  document.page = documentPage;
  document.images = [...drawingImages, ...images];
  document.headerImage = headerImage;
  document.footerPageNumbers = footerPageNumbers;
  await context.output?.documentMetadata(documentMetadataFromDoc(document));
  return { document, resources };
}
