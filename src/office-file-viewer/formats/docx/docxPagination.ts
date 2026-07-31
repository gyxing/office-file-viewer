import type {
  DocxBlock,
  DocxPageContent,
  DocxParagraphBlock,
  DocxTableBlock,
} from '../../services/docx/types';

/** 单个 DOCX 测量批次允许包含的普通内容块上限。 */
export const DOCX_MEASURE_BLOCK_LIMIT = 100;
/** 单个 DOCX 测量批次允许包含的表格行上限。 */
export const DOCX_MEASURE_TABLE_ROW_LIMIT = 200;

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
  /** 各表格行的测量高度。 */
  rowHeights?: readonly number[];
  /** 当前表格行相对表格顶部的偏移。 */
  rowOffset?: number;
  /** 当前表格在拆分前包含的总行数。 */
  originalTableRowCount?: number;
  /** 纯文本段落在浏览器排版后的各行 UTF-16 结束偏移。 */
  paragraphLineEndOffsets?: readonly number[];
  /** 与 paragraphLineEndOffsets 一一对应的行级流式高度。 */
  paragraphLineHeights?: readonly number[];
};

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
function sliceMeasuredParagraph(
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
  const contentHeight =
    sourcePage.page.minHeight -
    sourcePage.page.marginTop -
    sourcePage.page.marginBottom;
  if (
    measuredBlocks.reduce((sum, item) => sum + item.height, 0) <=
    contentHeight + 120
  ) {
    return [sourcePage];
  }

  const pages: DocxPageContent[] = [];
  let currentBlocks: DocxBlock[] = [];
  let currentHeight = 0;
  const pushPage = () => {
    if (!currentBlocks.length) return;
    pages.push({
      ...sourcePage,
      id: `${sourcePage.id}-flow-${pages.length + 1}`,
      blocks: currentBlocks,
    });
    currentBlocks = [];
    currentHeight = 0;
  };
  const appendBlock = (block: DocxBlock, height: number) => {
    if (currentBlocks.length && currentHeight + height > contentHeight + 1) {
      pushPage();
    }
    currentBlocks.push(block);
    currentHeight += height;
  };

  measuredBlocks.forEach((measurement) => {
    const { block } = measurement;
    const lineEndOffsets = measurement.paragraphLineEndOffsets;
    const lineHeights = measurement.paragraphLineHeights;
    if (
      block.type === 'paragraph' &&
      currentHeight + measurement.height > contentHeight + 1 &&
      canSplitMeasuredParagraph(block) &&
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
        if (
          currentBlocks.length &&
          currentHeight + lineHeights[lineStart] > contentHeight + 1
        ) {
          pushPage();
        }
        const availableHeight = contentHeight - currentHeight;
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
        const startOffset = lineStart === 0 ? 0 : lineEndOffsets[lineStart - 1];
        const endOffset = lineEndOffsets[lineEnd - 1];
        appendBlock(
          sliceMeasuredParagraph(block, startOffset, endOffset),
          fragmentHeight,
        );
        lineStart = lineEnd;
        if (lineStart < lineEndOffsets.length) pushPage();
      }
      return;
    }
    if (
      block.type !== 'table' ||
      (measurement.height <= contentHeight * 0.6 &&
        (measurement.originalTableRowCount ?? block.rows.length) <=
          block.rows.length)
    ) {
      appendBlock(block, measurement.height);
      return;
    }
    const rowHeights = measurement.rowHeights;
    if (!rowHeights?.length || rowHeights.length !== block.rows.length) {
      appendBlock(block, measurement.height);
      return;
    }
    let rowStart = 0;
    let rowHeight = 0;
    const appendRows = (rowEnd: number) => {
      if (rowEnd <= rowStart) return;
      const absoluteStart = (measurement.rowOffset ?? 0) + rowStart;
      const absoluteEnd = (measurement.rowOffset ?? 0) + rowEnd;
      appendBlock(
        {
          ...block,
          id: `${block.sourceBlockId ?? block.id}-rows-${
            absoluteStart + 1
          }-${absoluteEnd}`,
          sourceBlockId: block.sourceBlockId ?? block.id,
          rows: block.rows.slice(rowStart, rowEnd),
        },
        rowHeight,
      );
      rowStart = rowEnd;
      rowHeight = 0;
    };
    rowHeights.forEach((height, rowIndex) => {
      if (
        rowIndex > rowStart &&
        currentHeight + rowHeight + height > contentHeight + 1
      ) {
        appendRows(rowIndex);
        pushPage();
      }
      rowHeight += height;
    });
    appendRows(block.rows.length);
  });
  pushPage();
  return pages.length ? pages : [sourcePage];
}
