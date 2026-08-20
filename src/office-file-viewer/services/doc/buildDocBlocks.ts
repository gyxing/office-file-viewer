import { DocBlockStreamBuilder } from './DocBlockStreamBuilder';
import type {
  DocBlockBuildOptions,
  DocImageSegment,
  DocTableCellLayout,
  DocTextSegment,
} from './docParseTypes';
import { parseDocFieldHyperlink } from './parseDocHyperlinks';
import { nextDocNumberPrefix } from './parseDocNumbering';
import type {
  DocBlock,
  DocImage,
  DocListBlock,
  DocParagraphBlock,
  DocTableBlock,
  DocTableStyle,
  DocTextInline,
  DocTextStyle,
} from './types';
/** DOC 默认页面扣除左右页边距后的正文宽度。 */
const DEFAULT_DOC_CONTENT_WIDTH = 554;

/** DOC 无法解析字体信息时使用的默认字体回退栈。 */
const DOC_FONT_FAMILY =
  '"Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", Arial, sans-serif';
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
  /** 按当前文本执行正则匹配。 */
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

/** 合并基础样式与后续覆盖样式。 */
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

/** 用源文档样式覆盖推断样式，同时保留未声明字段。 */
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

/** DOC 域代码或域结果对应的文本片段。 */
type DocFieldTextChunk = {
  /** 文本内容。 */
  text: string;
  /** 当前文本块对应的源文档片段。 */
  segment: DocTextSegment;
};

/** 内部标记：EMBED 字段结果应消费 ObjectPool 预览，而非普通随文图片。 */
const DOC_OBJECT_IMAGE_MARKER = '\ue000';
/** 内部标记：脚注或尾注引用在标准化文本中保留独立行内节点。 */
const DOC_NOTE_REFERENCE_MARKER = '\ue001';

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
      if (!chunk.text && !chunk.segment.bookmarkMarkers?.length) return;
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
  const visibleFieldValue = (frame: FieldFrame) => {
    const instructionText = frame.instruction
      .map((chunk) => chunk.text)
      .join('');
    const hyperlink = parseDocFieldHyperlink(instructionText);
    const embeddedObject = /^\s*EMBED\b/i.test(instructionText);
    const visible = frame.result.length
      ? frame.result.map((chunk) => ({
          ...chunk,
          text: embeddedObject
            ? chunk.text.replace(/\u0001/g, DOC_OBJECT_IMAGE_MARKER)
            : chunk.text,
          segment: hyperlink ? { ...chunk.segment, hyperlink } : chunk.segment,
        }))
      : frame.instruction.flatMap((chunk) => {
          const anchors = chunk.text.match(/[\u0001\u0008]/g)?.join('') ?? '';
          return anchors ? [{ ...chunk, text: anchors }] : [];
        });
    const markers = [...frame.instruction, ...frame.result]
      .flatMap((chunk) => chunk.segment.bookmarkMarkers ?? [])
      .filter(
        (marker, index, items) =>
          items.findIndex((item) => item.markerId === marker.markerId) ===
          index,
      );
    if (!markers.length) return visible;
    const first = visible[0] ?? frame.result[0] ?? frame.instruction[0];
    if (!first) return visible;
    if (!visible.length) {
      return [
        {
          text: '',
          segment: { ...first.segment, bookmarkMarkers: markers },
        },
      ];
    }
    visible[0] = {
      ...visible[0],
      segment: { ...visible[0].segment, bookmarkMarkers: markers },
    };
    return visible;
  };

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
  images: DocImage[],
  drawingImages: Array<DocImage | undefined> = [],
  objectImages: DocImage[] = [],
) {
  let imageIndex = 0;
  let drawingImageIndex = 0;
  let objectImageIndex = 0;

  return preserveDocFieldResultSegments(segments).flatMap((segment) => {
    const normalizedText = segment.noteReference
      ? DOC_NOTE_REFERENCE_MARKER
      : normalizeDocText(segment.text);

    return normalizedText
      .split(/(\n|\f|\u0001|\u0008|\ue000|\ue001)/)
      .map((text, textIndex): DocImageSegment => {
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
            hyperlink: segment.hyperlink,
            review: segment.review,
            bookmarkMarkers:
              textIndex === 0 ? segment.bookmarkMarkers : undefined,
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
            hyperlink: segment.hyperlink,
            review: segment.review,
            bookmarkMarkers:
              textIndex === 0 ? segment.bookmarkMarkers : undefined,
          };
        }
        if (text === DOC_NOTE_REFERENCE_MARKER) {
          return {
            text,
            style: segment.style,
            noteReference: segment.noteReference,
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
            bookmarkMarkers:
              textIndex === 0 ? segment.bookmarkMarkers : undefined,
          };
        }
        if (text === DOC_OBJECT_IMAGE_MARKER) {
          const image = objectImages[objectImageIndex];
          if (image) objectImageIndex += 1;
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
            hyperlink: segment.hyperlink,
            review: segment.review,
            bookmarkMarkers:
              textIndex === 0 ? segment.bookmarkMarkers : undefined,
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
          hyperlink: segment.hyperlink,
          review: segment.review,
          bookmarkMarkers:
            textIndex === 0 ? segment.bookmarkMarkers : undefined,
        };
      })
      .filter(
        (item) =>
          item.image ||
          item.noteReference ||
          item.bookmarkMarkers?.length ||
          (item.text.length &&
            item.text !== '\u0001' &&
            item.text !== '\u0008' &&
            item.text !== DOC_OBJECT_IMAGE_MARKER &&
            item.text !== DOC_NOTE_REFERENCE_MARKER),
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

function sameInlineHyperlink(left: DocTextInline, right: DocTextInline) {
  if (left.type !== 'text' || right.type !== 'text') return false;
  return (
    JSON.stringify(left.hyperlink ?? null) ===
      JSON.stringify(right.hyperlink ?? null) &&
    JSON.stringify(left.review ?? null) === JSON.stringify(right.review ?? null)
  );
}

/** 合并 `mergeAdjacentInlines` 接收的多份数据。 */
function mergeAdjacentInlines(inlines: DocTextInline[]) {
  const merged: DocTextInline[] = [];

  inlines.forEach((inline) => {
    if (inline.type !== 'text') {
      merged.push({ ...inline });
      return;
    }
    if (!inline.text) return;
    const previous = merged[merged.length - 1];
    if (
      previous?.type === 'text' &&
      sameInlineStyle(previous.style, inline.style) &&
      sameInlineHyperlink(previous, inline)
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
        inline.type !== 'text' ||
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
    if (inline.type !== 'text') {
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
    if (inline.type === 'bookmark' || inline.type === 'note-reference') {
      result.push(inline);
      return;
    }
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
  const availableWidth = DEFAULT_DOC_CONTENT_WIDTH;
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
    (mergedStyle.fontSize ?? 0) >= 28 && mergedStyle.textAlign === 'center';
  if (
    shouldUseSourceLineMultiplier &&
    mergedStyle.lineHeightMultiplier !== undefined
  ) {
    // 大字号居中标题按字体倍数生成行框；目录仍保留源文档网格换算出的绝对行距。
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

/** 将带样式的 DOC 文本片段组装为段落、列表和表格块。 */
export async function buildDocBlocksFromSegments(
  segments: DocTextSegment[],
  images: DocImage[],
  options: DocBlockBuildOptions,
  drawingImages?: Array<DocImage | undefined>,
  objectImages?: DocImage[],
): Promise<DocBlock[]> {
  const pendingTableRows: PendingTableRow[] = [];
  const pendingTableCells: PendingTableCell[] = [];
  const pendingListItems: ParsedListLine[] = [];
  const normalizedSegments = normalizeDocTextSegments(
    segments,
    images,
    drawingImages,
    objectImages,
  );
  const referencedBodyImageIds = new Set(
    normalizedSegments
      .map((segment) => segment.image?.id)
      .filter((imageId): imageId is string => Boolean(imageId)),
  );
  const imagesWithoutTextAnchor = images.filter(
    (image) => !referencedBodyImageIds.has(image.id),
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
    const shouldUseCompactTocLineHeight =
      isTableOfContents && tocSpacingAfter !== undefined && tocSpacingAfter > 0;
    const style =
      isTableOfContents &&
      sourceStyle?.useDocumentGrid !== false &&
      options.defaultGridLineHeight
        ? {
            ...sourceStyle,
            // 可在单页容纳的长目录以原段落行距配合均摊段后距；
            // 真正跨页的目录继续使用文档网格，避免统一压缩破坏源分页。
            lineHeight: shouldUseCompactTocLineHeight
              ? explicitLineHeight ?? sourceStyle?.lineHeight ?? 1.5
              : Math.max(
                  explicitLineHeight ?? 0,
                  options.defaultGridLineHeight,
                ),
          }
        : !inTable &&
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
    // 段落未显式脱离文档网格时保留该语义，渲染器才能复原 Word 的字符节距。
    const resolvedStyle =
      !inTable &&
      options.defaultGridLineHeight !== undefined &&
      sourceStyle?.useDocumentGrid !== false
        ? { ...style, useDocumentGrid: true }
        : style;
    return {
      text,
      inlines: mergeAdjacentInlines(trimInlines(currentLineInlines)),
      style: resolvedStyle,
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
      if (line.inlines.some((inline) => inline.type !== 'text')) {
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
    for (const marker of segment.bookmarkMarkers ?? []) {
      currentLineInlines.push({ type: 'bookmark', ...marker });
    }
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
          hyperlink: segment.hyperlink,
          review: segment.review,
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
    } else if (segment.noteReference) {
      currentLineInlines.push({
        type: 'note-reference',
        ...segment.noteReference,
        style: segment.style,
      });
      currentLineSegments.push(segment);
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
        hyperlink: segment.hyperlink,
        review: segment.review,
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
  // 部分旧 DOC/WPS 只保留少量 0x01 图片锚点，但图片流仍按正文顺序保存完整内容。
  // 将剩余图片继续放回正文流，分页器才能按尺寸恢复后续图片页，而不是在尾页降级成 Gallery。
  for (const image of imagesWithoutTextAnchor) {
    await builder.add(
      createParagraphBlock('', builder.nextSourceIndex, [
        { type: 'image', image },
      ]),
    );
  }
  return builder.finish();
}

/** 将纯文本降级内容组装为 DOC 块模型。 */
export async function buildDocBlocksFromText(
  text: string,
  options: DocBlockBuildOptions,
): Promise<DocBlock[]> {
  return buildDocBlocksFromSegments([{ text }], [], options);
}
