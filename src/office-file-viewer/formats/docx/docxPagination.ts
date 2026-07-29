import type {
  DocxBlock,
  DocxPageContent,
  DocxTableBlock,
} from '../../services/docx/types';

export const DOCX_MEASURE_BLOCK_LIMIT = 100;
export const DOCX_MEASURE_TABLE_ROW_LIMIT = 200;

export type DocxMeasurementContext = {
  previousBlock?: DocxBlock;
  sectionIndex: number;
  continuedTable?: {
    sourceBlockId: string;
    consumedRowCount: number;
  };
};

export type DocxMeasurementBatch = {
  id: string;
  revision: number;
  sourcePage: DocxPageContent;
  blocks: readonly DocxBlock[];
  contextBefore?: DocxMeasurementContext;
  rowOffsets: Readonly<Record<string, number>>;
  originalTableRowCounts: Readonly<Record<string, number>>;
  endOfSourcePage: boolean;
};

export type DocxMeasuredBlock = {
  block: DocxBlock;
  height: number;
  rowHeights?: readonly number[];
  rowOffset?: number;
  originalTableRowCount?: number;
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
