import type {
  DocBinaryContent,
  DocCharacterRun,
  DocDrawingTextBoxRange,
  DocFib,
  DocParagraphRun,
  DocPiece,
  DocSectionLayout,
  DocTableCellLayout,
  DocTextSegment,
} from './docParseTypes';
import { readDocNumberingCatalog } from './parseDocNumbering';
import {
  isDocParagraphTocStyle,
  parseDocStyleOutlineCatalog,
  readDocParagraphOutlineLevel,
  readDocParagraphStyleChain,
  type DocStyleOutlineCatalog,
} from './parseDocStyleOutline';
import { readDocBookmarks } from './readDocBookmarks';
import type { DocPage, DocTableBlock, DocTextStyle } from './types';

function mergeBinaryTextStyle(
  base: DocTextStyle | undefined,
  next: DocTextStyle | undefined,
): DocTextStyle | undefined {
  if (!base && !next) return undefined;
  return { ...base, ...next };
}

/** 按字体编号索引的 DOC 字体族名称。 */
type DocFontTable = string[];

// 旧版 .doc 是 OLE/CFB 二进制容器，不是 zip；这里实现最小可用的前端降级解析。
/** DOC 缺少页面设置时使用的默认页面尺寸和边距。 */
export const DEFAULT_DOC_PAGE: DocPage = {
  width: 794,
  minHeight: 1123,
  marginTop: 96,
  marginRight: 120,
  marginBottom: 96,
  marginLeft: 120,
};

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

function readUint16(view: DataView, offset: number) {
  return view.getUint16(offset, true);
}

function readUint32(view: DataView, offset: number) {
  return view.getUint32(offset, true);
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

/** 读取 FIB 并确定表流、story 字符数和各结构偏移。 */
export function readDocFib(wordDocument: Uint8Array): DocFib {
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
    fcSttbfBkmk: readFibField(wordDocument, 322),
    lcbSttbfBkmk: readFibField(wordDocument, 326),
    fcPlcfBkf: readFibField(wordDocument, 330),
    lcbPlcfBkf: readFibField(wordDocument, 334),
    fcPlcfBkl: readFibField(wordDocument, 338),
    lcbPlcfBkl: readFibField(wordDocument, 342),
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

/** 从 PlcftxbxTxt 读取每个 OfficeArt 文本框在文本框 story 中的字符范围。 */
function readDocDrawingTextBoxRanges(
  tableStream: Uint8Array,
  fib: DocFib,
): Array<DocDrawingTextBoxRange | undefined> {
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
    (_, index): DocDrawingTextBoxRange | undefined => {
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
      return mergeBinaryTextStyle(
        mergeBinaryTextStyle(style, paragraphStyle),
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
    const style = mergeBinaryTextStyle(
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
      (style, run) => mergeBinaryTextStyle(style, run.style),
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
      (style, run) => mergeBinaryTextStyle(style, run.style),
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
            mergeBinaryTextStyle(
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
    charStart: piece.charStart,
    charEnd: piece.charEnd,
    style,
  };
}

/** 一次建立 DOC 二进制正文后续解析所需的结构索引。 */
export function readDocBinaryContent(input: {
  /** WordDocument 主流。 */
  wordDocument: Uint8Array;
  /** FIB 指定的 0Table 或 1Table 流。 */
  tableStream: Uint8Array;
  /** PAPX 外置属性可能引用的 Data 流。 */
  dataStream?: Uint8Array;
  /** 预先读取的 DOC 文件信息块。 */
  fib: DocFib;
}): DocBinaryContent {
  const { wordDocument, tableStream, dataStream, fib } = input;
  const pieces = parsePieces(tableStream, fib);
  if (!pieces.length) {
    throw new Error('暂未能识别该 DOC 文件的正文片段表');
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
  const bookmarkResult = readDocBookmarks(tableStream, fib);
  const sections = readDocSections(wordDocument, tableStream, fib);
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

  return {
    fib,
    pieces,
    sections,
    characterRuns,
    paragraphRuns,
    drawingTextBoxRanges: readDocDrawingTextBoxRanges(tableStream, fib),
    normalStyle,
    outlineCatalog,
    numbering,
    bookmarks: bookmarkResult.bookmarks,
    bookmarkWarnings: bookmarkResult.warnings,
  };
}

/** 读取指定 DOC story 字符区间并保留片段样式与结构属性。 */
export function readDocStorySegments(
  wordDocument: Uint8Array,
  content: DocBinaryContent,
  charStart: number,
  charEnd: number,
): DocTextSegment[] {
  const pieces = slicePiecesByCharacterRange(
    content.pieces,
    charStart,
    charEnd,
  );
  return textSegmentsFromPieces(
    wordDocument,
    pieces,
    content.characterRuns,
    content.paragraphRuns,
  );
}
