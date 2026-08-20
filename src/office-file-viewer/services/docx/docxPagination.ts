import { attachDocxFootnotesToPage } from './docxNoteReferences';
import type {
  DocxBlock,
  DocxNote,
  DocxPageContent,
  DocxParagraphBlock,
  DocxTableBlock,
} from './types';

/** 单个 DOCX 测量批次允许包含的普通内容块上限。 */
export const DOCX_MEASURE_BLOCK_LIMIT = 100;
/** 单个 DOCX 测量批次允许包含的表格行上限。 */
export const DOCX_MEASURE_TABLE_ROW_LIMIT = 200;
/** 仅忽略浏览器布局取整造成的亚像素误差，避免真实溢出内容被纸张裁切。 */
const DOCX_PAGE_OVERFLOW_TOLERANCE = 1;
/** 脚注分隔线及其上下间距需要占用的最小高度。 */
const DOCX_FOOTNOTE_SEPARATOR_HEIGHT = 8;

/** DOCX 分页测量过程需要保留的上下文。 */
export type DocxMeasurementContext = {
  /** 当前内容块之前的相邻内容块。 */
  previousBlock?: DocxBlock;
  /** 文档节的零基索引。 */
  sectionIndex: number;
  /** 当前批次是否延续上一页的表格。 */
  continuedTable?: {
    /** 分页拆分前源内容块的稳定标识。 */
    sourceBlockId: string;
    /** 当前批次已经消费的表格行数。 */
    consumedRowCount: number;
  };
};

/** 一次提交到隐藏容器测量的 DOCX 内容批次。 */
export type DocxMeasurementBatch = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 数据源变更时递增的修订号。 */
  revision: number;
  /** 分页处理前所属的源页面。 */
  sourcePage: DocxPageContent;
  /** 按源文档顺序排列的内容块。 */
  blocks: readonly DocxBlock[];
  /** 当前批次开始前保留的分页上下文。 */
  contextBefore?: DocxMeasurementContext;
  /** 各表格行相对表格顶部的偏移。 */
  rowOffsets: Readonly<Record<string, number>>;
  /** 各源表格在拆分前包含的行数。 */
  originalTableRowCounts: Readonly<Record<string, number>>;
  /** 当前批次是否到达源页面末尾。 */
  endOfSourcePage: boolean;
};

/** 完成高度测量的 DOCX 内容块。 */
export type DocxMeasuredBlock = {
  /** 当前负责处理或渲染的内容块。 */
  block: DocxBlock;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 内容块成为页首时需要额外占用的上边距。 */
  leadingSpacing: number;
  /** 内容块位于页尾时最后一个可见字形占用的高度。 */
  pageEndHeight?: number;
  /** 各表格行的测量高度。 */
  rowHeights?: readonly number[];
  /** 各表格行可在不裁切可见内容时使用的纵向拆分位置。 */
  rowBreakOffsets?: readonly (readonly number[])[];
  /** 当前表格行相对表格顶部的偏移。 */
  rowOffset?: number;
  /** 当前表格在拆分前包含的总行数。 */
  originalTableRowCount?: number;
  /** 纯文本段落在浏览器排版后的各行 UTF-16 结束偏移。 */
  paragraphLineEndOffsets?: readonly number[];
  /** 与 paragraphLineEndOffsets 一一对应的行级流式高度。 */
  paragraphLineHeights?: readonly number[];
  /** 当前正文块首次引用的脚注正文实测高度。 */
  footnoteReserveHeight?: number;
  /** 当前正文块首次引用的脚注分页片段。 */
  measuredFootnotes?: readonly DocxMeasuredFootnote[];
};

/** 已按页面容量拆分的单条脚注。 */
export type DocxMeasuredFootnote = Readonly<{
  /** 源脚注标识。 */
  noteId: string;
  /** 首段显示在引用页，其余片段按顺序进入脚注续页。 */
  fragments: readonly Readonly<{
    /** 当前页使用的脚注块片段。 */
    note: DocxNote;
    /** 当前片段实测高度。 */
    height: number;
  }>[];
}>;

function splitTableForMeasurement(
  table: DocxTableBlock,
  rowStart: number,
  rowEnd: number,
) {
  return {
    ...table,
    id: `${table.id}-measure-${rowStart + 1}-${rowEnd}`,
    sourceBlockId: table.sourceBlockId ?? table.id,
    rows: table.rows.slice(rowStart, rowEnd),
  };
}

/** 判断段落是否可在不破坏复杂行内对象的前提下按浏览器实测行拆分。 */
export function canSplitMeasuredParagraph(block: DocxParagraphBlock) {
  return (
    !block.position &&
    !block.isTableOfContents &&
    !block.backgroundColor &&
    !block.borderTop &&
    !block.borderRight &&
    !block.borderBottom &&
    !block.borderLeft &&
    !block.paddingTop &&
    !block.paddingRight &&
    !block.paddingBottom &&
    !block.paddingLeft &&
    block.inlines.length > 0 &&
    block.inlines.every((inline) => inline.type === 'text')
  );
}

/** 按 UTF-16 偏移裁剪纯文本段落，同时保留跨行的文字样式。 */
export function sliceMeasuredParagraph(
  block: DocxParagraphBlock,
  startOffset: number,
  endOffset: number,
) {
  let inlineOffset = 0;
  const inlines = block.inlines.flatMap((inline) => {
    if (inline.type !== 'text') return [];
    const inlineStart = inlineOffset;
    const inlineEnd = inlineStart + inline.text.length;
    inlineOffset = inlineEnd;
    const sliceStart = Math.max(startOffset, inlineStart) - inlineStart;
    const sliceEnd = Math.min(endOffset, inlineEnd) - inlineStart;
    if (sliceStart >= sliceEnd) return [];
    return [{ ...inline, text: inline.text.slice(sliceStart, sliceEnd) }];
  });
  const isFirstFragment = startOffset === 0;
  const isLastFragment = endOffset === inlineOffset;
  return {
    ...block,
    id: `${block.sourceBlockId ?? block.id}-chars-${
      startOffset + 1
    }-${endOffset}`,
    sourceBlockId: block.sourceBlockId ?? block.id,
    inlines,
    text: inlines.map((inline) => inline.text).join(''),
    outlineLevel: isFirstFragment ? block.outlineLevel : undefined,
    keepNext: isLastFragment ? block.keepNext : undefined,
    spacingBefore: isFirstFragment ? block.spacingBefore : 0,
    spacingAfter: isLastFragment ? block.spacingAfter : 0,
    firstLineIndent: isFirstFragment ? block.firstLineIndent : 0,
  };
}

/** 按正文块和表格行双阈值切分隐藏测量批次。 */
export function createDocxMeasurementBatches(
  sourcePage: DocxPageContent,
  firstRevision: number,
): DocxMeasurementBatch[] {
  // 整页定位画布已对应源文档物理页，交给流式测量会把画布内部内容误拆成多页。
  if (sourcePage.preservePhysicalPage) return [];

  const batches: DocxMeasurementBatch[] = [];
  let blocks: DocxBlock[] = [];
  let rowCount = 0;
  let previousBlock: DocxBlock | undefined;
  let rowOffsets: Record<string, number> = {};
  let originalTableRowCounts: Record<string, number> = {};

  const push = () => {
    if (!blocks.length) return;
    batches.push({
      id: `${sourcePage.id}-batch-${batches.length + 1}`,
      revision: firstRevision + batches.length,
      sourcePage,
      blocks,
      contextBefore: previousBlock
        ? { previousBlock, sectionIndex: 0 }
        : undefined,
      rowOffsets,
      originalTableRowCounts,
      endOfSourcePage: false,
    });
    previousBlock = blocks[blocks.length - 1];
    blocks = [];
    rowCount = 0;
    rowOffsets = {};
    originalTableRowCounts = {};
  };

  sourcePage.blocks.forEach((block) => {
    if (block.type !== 'table') {
      if (blocks.length >= DOCX_MEASURE_BLOCK_LIMIT) push();
      blocks.push(block);
      return;
    }
    let rowStart = 0;
    while (rowStart < block.rows.length) {
      if (
        blocks.length >= DOCX_MEASURE_BLOCK_LIMIT ||
        rowCount >= DOCX_MEASURE_TABLE_ROW_LIMIT
      ) {
        push();
      }
      const availableRows = DOCX_MEASURE_TABLE_ROW_LIMIT - rowCount;
      const rowEnd = Math.min(
        block.rows.length,
        rowStart + Math.max(1, availableRows),
      );
      const tablePart =
        rowStart === 0 && rowEnd === block.rows.length
          ? block
          : splitTableForMeasurement(block, rowStart, rowEnd);
      blocks.push(tablePart);
      rowOffsets[tablePart.id] = rowStart;
      originalTableRowCounts[tablePart.id] = block.rows.length;
      rowCount += rowEnd - rowStart;
      rowStart = rowEnd;
      if (rowStart < block.rows.length) push();
    }
  });
  push();
  if (batches.length) {
    batches[batches.length - 1] = {
      ...batches[batches.length - 1],
      endOfSourcePage: true,
    };
  }
  return batches;
}

/** 使用浏览器测得的块和行高，复用现有 DOCX 溢出分页规则。 */
export function paginateMeasuredDocxPage(
  sourcePage: DocxPageContent,
  measuredBlocks: readonly DocxMeasuredBlock[],
) {
  // 同时保护同步测量路径，确保整页画布在不同数据规模下都保持一个物理页。
  if (sourcePage.preservePhysicalPage) return [sourcePage];

  const contentHeight =
    sourcePage.page.minHeight -
    sourcePage.page.marginTop -
    sourcePage.page.marginBottom;
  // 源页面首段保留自身段前距；仅自动续页会抑制新页首块的普通段前距。
  const firstPageLeadingSpacing = measuredBlocks[0]?.leadingSpacing ?? 0;
  const measuredContentHeight = measuredBlocks.reduce(
    (sum, item) => sum + item.height,
    firstPageLeadingSpacing,
  );
  const measuredFootnoteHeight = Math.min(
    Math.max(
      0,
      measuredBlocks.reduce(
        (height, item) => height + (item.footnoteReserveHeight ?? 0),
        0,
      ),
    ),
    Math.max(0, contentHeight - 24),
  );
  const measuredFootnotes = new Map(
    measuredBlocks
      .flatMap((item) => item.measuredFootnotes ?? [])
      .map((note) => [note.noteId, note] as const),
  );
  const appendFootnoteContinuations = (
    sourcePages: readonly DocxPageContent[],
  ) => {
    const result: DocxPageContent[] = [];
    sourcePages.forEach((page) => {
      const pageNotes = page.footnotes ?? [];
      const firstFragments = pageNotes.map(
        (note) => measuredFootnotes.get(note.id)?.fragments[0]?.note ?? note,
      );
      result.push({
        ...page,
        footnotes: firstFragments.length ? firstFragments : undefined,
      });
      const continuations = pageNotes.flatMap(
        (note) => measuredFootnotes.get(note.id)?.fragments.slice(1) ?? [],
      );
      let continuationNotes: DocxNote[] = [];
      let continuationHeight = 0;
      const pushContinuation = () => {
        if (!continuationNotes.length) return;
        result.push({
          ...page,
          id: `${page.id}-footnote-continuation-${result.length + 1}`,
          preservePhysicalPage: false,
          blocks: [],
          footnotes: continuationNotes,
        });
        continuationNotes = [];
        continuationHeight = 0;
      };
      continuations.forEach((fragment) => {
        if (
          continuationNotes.length &&
          continuationHeight + fragment.height > contentHeight - 16
        ) {
          pushContinuation();
        }
        continuationNotes.push(fragment.note);
        continuationHeight += fragment.height;
      });
      pushContinuation();
    });
    return result;
  };
  if (
    measuredContentHeight +
      (measuredFootnoteHeight > 0
        ? measuredFootnoteHeight + DOCX_FOOTNOTE_SEPARATOR_HEIGHT
        : 0) <=
    contentHeight + DOCX_PAGE_OVERFLOW_TOLERANCE
  ) {
    return appendFootnoteContinuations([sourcePage]);
  }

  const pages: DocxPageContent[] = [];
  let currentBlocks: DocxBlock[] = [];
  let currentHeight = firstPageLeadingSpacing;
  let currentFootnoteHeight = 0;
  const normalizeFootnoteHeight = (height: number) =>
    Math.min(Math.max(0, height), Math.max(0, contentHeight - 24));
  const reservedFootnoteArea = (height: number) =>
    height > 0 ? height + DOCX_FOOTNOTE_SEPARATOR_HEIGHT : 0;
  const remainingBodyHeight = (additionalFootnoteHeight = 0) =>
    contentHeight -
    currentHeight -
    reservedFootnoteArea(
      normalizeFootnoteHeight(
        currentFootnoteHeight + Math.max(0, additionalFootnoteHeight),
      ),
    );
  const pushPage = () => {
    if (!currentBlocks.length) return;
    pages.push(
      attachDocxFootnotesToPage(
        {
          ...sourcePage,
          id: `${sourcePage.id}-flow-${pages.length + 1}`,
          blocks: currentBlocks,
          footnotes: undefined,
        },
        sourcePage.footnotes ?? [],
      ),
    );
    currentBlocks = [];
    currentHeight = 0;
    currentFootnoteHeight = 0;
  };
  const appendBlock = (
    block: DocxBlock,
    height: number,
    _leadingSpacing = 0,
    pageEndHeight = height,
    footnoteReserveHeight = 0,
  ) => {
    const nextFootnoteHeight = normalizeFootnoteHeight(
      currentFootnoteHeight + footnoteReserveHeight,
    );
    if (
      currentBlocks.length &&
      currentHeight + pageEndHeight + reservedFootnoteArea(nextFootnoteHeight) >
        contentHeight + DOCX_PAGE_OVERFLOW_TOLERANCE
    ) {
      pushPage();
    }
    currentBlocks.push(block);
    // Word 会在自动分页后的页首抑制普通段前距，页内间距已包含在相邻块测量高度中。
    currentHeight += height;
    currentFootnoteHeight = normalizeFootnoteHeight(
      currentFootnoteHeight + footnoteReserveHeight,
    );
  };

  const getKeepWithNextTerminalHeight = (item: DocxMeasuredBlock) => {
    if (item.block.type === 'paragraph') {
      if (
        item.block.widowControl !== false &&
        !item.block.keepLines &&
        item.paragraphLineHeights?.length
      ) {
        // 默认孤行控制只要求标题后的正文至少保留两行。
        return item.paragraphLineHeights
          .slice(0, Math.min(2, item.paragraphLineHeights.length))
          .reduce((height, lineHeight) => height + lineHeight, 0);
      }
      // 关闭孤行控制时，WPS 会把可完整放入新页的下一段与标题整段成组。
      return item.height;
    }
    if (item.block.type === 'table' && item.rowHeights?.length) {
      return item.rowHeights[0];
    }
    return item.pageEndHeight ?? item.height;
  };

  const getKeepWithNextHeight = (startIndex: number) => {
    let height = 0;
    for (let index = startIndex; index < measuredBlocks.length; index += 1) {
      const item = measuredBlocks[index];
      if (item.block.type !== 'paragraph' || !item.block.keepNext) {
        // 按终端正文的孤行设置或表格首行计算 keepNext 成组高度。
        height += getKeepWithNextTerminalHeight(item);
        break;
      }
      height += item.height;
    }
    return height;
  };

  measuredBlocks.forEach((measurement, measurementIndex) => {
    const { block } = measurement;

    if (block.type === 'paragraph' && block.keepNext && currentBlocks.length) {
      const groupedHeight = getKeepWithNextHeight(measurementIndex);
      if (
        groupedHeight <= contentHeight + DOCX_PAGE_OVERFLOW_TOLERANCE &&
        currentHeight + groupedHeight >
          contentHeight + DOCX_PAGE_OVERFLOW_TOLERANCE
      ) {
        pushPage();
      }
    }
    const lineEndOffsets = measurement.paragraphLineEndOffsets;
    const lineHeights = measurement.paragraphLineHeights;
    if (
      block.type === 'paragraph' &&
      currentHeight +
        (measurement.pageEndHeight ?? measurement.height) +
        reservedFootnoteArea(
          normalizeFootnoteHeight(
            currentFootnoteHeight + (measurement.footnoteReserveHeight ?? 0),
          ),
        ) >
        contentHeight + DOCX_PAGE_OVERFLOW_TOLERANCE &&
      canSplitMeasuredParagraph(block) &&
      (!block.keepLines || measurement.height > contentHeight) &&
      lineEndOffsets &&
      lineHeights &&
      lineEndOffsets.length > 1 &&
      lineEndOffsets.length === lineHeights.length &&
      lineEndOffsets[lineEndOffsets.length - 1] ===
        block.inlines.reduce(
          (textLength, inline) =>
            textLength + (inline.type === 'text' ? inline.text.length : 0),
          0,
        )
    ) {
      let lineStart = 0;
      while (lineStart < lineEndOffsets.length) {
        const fragmentLeadingSpacing =
          lineStart === 0 ? measurement.leadingSpacing : 0;
        if (
          currentBlocks.length &&
          currentHeight + lineHeights[lineStart] >
            contentHeight + DOCX_PAGE_OVERFLOW_TOLERANCE
        ) {
          pushPage();
        }
        const availableHeight = remainingBodyHeight();
        let lineEnd = lineStart;
        let fragmentHeight = 0;
        while (
          lineEnd < lineEndOffsets.length &&
          (lineEnd === lineStart ||
            fragmentHeight + lineHeights[lineEnd] <= availableHeight + 1)
        ) {
          fragmentHeight += lineHeights[lineEnd];
          lineEnd += 1;
        }
        if (block.widowControl !== false && lineEnd < lineEndOffsets.length) {
          const fragmentLineCount = lineEnd - lineStart;
          const remainingLineCount = lineEndOffsets.length - lineEnd;
          if (
            lineStart === 0 &&
            currentBlocks.length > 0 &&
            (fragmentLineCount < 2 ||
              (remainingLineCount === 1 && fragmentLineCount <= 2))
          ) {
            pushPage();
            continue;
          }
          if (remainingLineCount === 1 && fragmentLineCount > 2) {
            lineEnd -= 1;
            fragmentHeight -= lineHeights[lineEnd];
          }
        }
        const startOffset = lineStart === 0 ? 0 : lineEndOffsets[lineStart - 1];
        const endOffset = lineEndOffsets[lineEnd - 1];
        appendBlock(
          sliceMeasuredParagraph(block, startOffset, endOffset),
          fragmentHeight,
          fragmentLeadingSpacing,
        );
        lineStart = lineEnd;
        if (lineStart < lineEndOffsets.length) pushPage();
      }
      return;
    }
    const availableHeight = remainingBodyHeight(
      measurement.footnoteReserveHeight,
    );
    if (
      block.type !== 'table' ||
      measurement.height <= availableHeight + DOCX_PAGE_OVERFLOW_TOLERANCE
    ) {
      appendBlock(
        block,
        measurement.height,
        measurement.leadingSpacing,
        measurement.pageEndHeight,
        measurement.footnoteReserveHeight,
      );
      return;
    }
    const rowHeights = measurement.rowHeights;
    if (!rowHeights?.length || rowHeights.length !== block.rows.length) {
      appendBlock(
        block,
        measurement.height,
        measurement.leadingSpacing,
        measurement.height,
        measurement.footnoteReserveHeight,
      );
      return;
    }
    const rowBreakOffsets = measurement.rowBreakOffsets;
    const trailingSpacing = Math.max(
      0,
      measurement.height - rowHeights.reduce((sum, height) => sum + height, 0),
    );
    let pendingRows: DocxTableBlock['rows'] = [];
    let pendingHeight = 0;
    let partIndex = 0;
    let hasPreviousRows = (measurement.rowOffset ?? 0) > 0;
    let pendingFootnoteHeight = measurement.footnoteReserveHeight ?? 0;
    const appendRows = (isFinalPart = false) => {
      if (!pendingRows.length) return;
      partIndex += 1;
      appendBlock(
        {
          ...block,
          id: `${block.sourceBlockId ?? block.id}-part-${
            (measurement.rowOffset ?? 0) + 1
          }-${partIndex}`,
          sourceBlockId: block.sourceBlockId ?? block.id,
          rows: pendingRows,
        },
        pendingHeight,
        hasPreviousRows ? 0 : measurement.leadingSpacing,
        pendingHeight,
        pendingFootnoteHeight,
      );
      pendingRows = [];
      pendingHeight = 0;
      hasPreviousRows = true;
      pendingFootnoteHeight = 0;
      if (isFinalPart) {
        // 表格拆分时行高之和不包含表格外框及与下一块折叠后的间距，末段需补回分页记账。
        currentHeight += trailingSpacing;
      }
    };
    block.rows.forEach((row, rowIndex) => {
      const sourceHeight = rowHeights[rowIndex];
      const safeBreaks = rowBreakOffsets?.[rowIndex] ?? [];
      const canSplitRow =
        !row.cantSplit &&
        safeBreaks.length > 0 &&
        row.cells.every((cell) => !cell.rowSpan || cell.rowSpan === 1);
      let fragmentOffset = 0;
      while (fragmentOffset < sourceHeight - 0.5) {
        const remainingHeight = sourceHeight - fragmentOffset;
        const availableHeight =
          remainingBodyHeight(pendingFootnoteHeight) - pendingHeight;
        if (remainingHeight <= availableHeight + DOCX_PAGE_OVERFLOW_TOLERANCE) {
          pendingRows.push(
            fragmentOffset > 0
              ? {
                  ...row,
                  id: `${row.id}-fragment-${Math.round(fragmentOffset)}`,
                  fragment: {
                    height: remainingHeight,
                    offset: fragmentOffset,
                    sourceHeight,
                  },
                }
              : row,
          );
          pendingHeight += remainingHeight;
          break;
        }

        // 固化本轮分片边界，避免同步查找闭包读取随后会更新的偏移量。
        const minimumBreakOffset = fragmentOffset + 4;
        const maximumBreakOffset =
          fragmentOffset + availableHeight + DOCX_PAGE_OVERFLOW_TOLERANCE;
        const breakOffset = canSplitRow
          ? [...safeBreaks]
              .reverse()
              .find(
                (offset) =>
                  offset > minimumBreakOffset &&
                  offset < sourceHeight - 4 &&
                  offset <= maximumBreakOffset,
              )
          : undefined;
        if (breakOffset !== undefined) {
          const fragmentHeight = breakOffset - fragmentOffset;
          pendingRows.push({
            ...row,
            id: `${row.id}-fragment-${Math.round(fragmentOffset)}-${Math.round(
              breakOffset,
            )}`,
            fragment: {
              height: fragmentHeight,
              offset: fragmentOffset,
              sourceHeight,
            },
          });
          pendingHeight += fragmentHeight;
          appendRows();
          pushPage();
          fragmentOffset = breakOffset;
          continue;
        }

        if (pendingRows.length) appendRows();
        if (currentBlocks.length) {
          pushPage();
          continue;
        }

        // 空白页仍无法找到安全拆点时保留整行，避免裁切图片或复杂单元格内容。
        pendingRows.push(row);
        pendingHeight += sourceHeight;
        break;
      }
    });
    appendRows(true);
  });
  pushPage();
  return appendFootnoteContinuations(pages.length ? pages : [sourcePage]);
}
