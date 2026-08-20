import { WordRevisionRecordCollector } from '../word/review/WordRevisionRecordCollector';
import type {
  DocxBlock,
  DocxInline,
  DocxPageContent,
  DocxParagraphBlock,
} from './types';

/** 收集单个段落及其形状内部的修订定位。 */
function collectParagraphRevisions(
  block: DocxParagraphBlock,
  collector: WordRevisionRecordCollector,
  pageIndex?: number,
) {
  const blockId = block.sourceBlockId ?? block.id;
  block.review?.revisions?.forEach((revision) => {
    collector.add(
      revision,
      {
        kind: 'word-range',
        blockId,
        startOffset: 0,
        endOffset: 0,
        pageIndex,
      },
      block.text,
    );
  });
  let offset = 0;
  const collectInline = (inline: DocxInline) => {
    const fragment = inline.type === 'text' ? inline.text : '';
    const length =
      inline.type === 'text'
        ? inline.text.length
        : inline.type === 'tab' || inline.type === 'note-reference'
        ? 1
        : 0;
    inline.review?.revisions?.forEach((revision) => {
      collector.add(
        revision,
        {
          kind: 'word-range',
          blockId,
          startOffset: offset,
          endOffset: offset + length,
          pageIndex,
        },
        fragment,
      );
    });
    offset += length;
    if (inline.type !== 'shape') return;
    inline.shape.items.forEach((item) => {
      (item.blocks ?? item.paragraphs ?? []).forEach((child) =>
        collectDocxBlockRevisions(child, collector, pageIndex),
      );
    });
  };
  block.inlines.forEach(collectInline);
}

/** 递归收集段落、表格和形状中的修订。 */
function collectDocxBlockRevisions(
  block: DocxBlock,
  collector: WordRevisionRecordCollector,
  pageIndex?: number,
) {
  if (block.type === 'paragraph') {
    collectParagraphRevisions(block, collector, pageIndex);
    return;
  }
  if (block.type !== 'table') return;
  block.rows.forEach((row) =>
    row.cells.forEach((cell) =>
      cell.blocks.forEach((child) =>
        collectDocxBlockRevisions(child, collector, pageIndex),
      ),
    ),
  );
}

/** 把一组 DOCX 正文块追加到共享修订记录收集器。 */
export function collectDocxRevisionRecords(
  blocks: readonly DocxBlock[],
  collector = new WordRevisionRecordCollector(),
  pageIndex?: number,
) {
  blocks.forEach((block) =>
    collectDocxBlockRevisions(block, collector, pageIndex),
  );
  return collector;
}

/** 收集页面正文及页眉变体中的全部修订。 */
export function collectDocxPageRevisionRecords(
  pages: readonly DocxPageContent[],
  collector = new WordRevisionRecordCollector(),
  pageIndexOffset = 0,
) {
  pages.forEach((page, index) => {
    const pageIndex = pageIndexOffset + index;
    collectDocxRevisionRecords(page.blocks, collector, pageIndex);
    const regions = page.headers;
    if (!regions) return;
    [regions.default, regions.first, regions.even].forEach((blocks) => {
      if (blocks) collectDocxRevisionRecords(blocks, collector, pageIndex);
    });
  });
  return collector;
}
