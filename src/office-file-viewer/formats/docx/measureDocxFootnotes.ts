import type { DocxMeasuredFootnote } from '../../services/docx/docxPagination';
import { sliceMeasuredParagraph } from '../../services/docx/docxPagination';
import type { DocxBlock, DocxNote, DocxPage } from '../../services/docx/types';
import { measureDocxParagraphLines } from './measureDocxParagraphLines';

function createMeasuredFootnote(
  note: DocxNote,
  item: HTMLElement,
  page: DocxPage,
): DocxMeasuredFootnote {
  const elementsById = new Map(
    Array.from(
      item.querySelectorAll<HTMLElement>('[data-office-word-block-id]'),
    ).map((element) => [element.dataset.officeWordBlockId ?? '', element]),
  );
  const measurements = note.blocks.map((block) => {
    const element = elementsById.get(block.id);
    const height = element?.offsetHeight ?? 0;
    return {
      block,
      height,
      ...(element ? measureDocxParagraphLines(element, block, height) : {}),
    };
  });
  const contentHeight = page.minHeight - page.marginTop - page.marginBottom;
  const firstCapacity = Math.max(120, contentHeight * 0.38);
  const continuationCapacity = Math.max(120, contentHeight - 24);
  const fragments: Array<{ note: DocxNote; height: number }> = [];
  let blocks: DocxBlock[] = [];
  let height = 0;
  const capacity = () =>
    fragments.length ? continuationCapacity : firstCapacity;
  const push = () => {
    if (!blocks.length) return;
    fragments.push({ note: { ...note, blocks }, height });
    blocks = [];
    height = 0;
  };

  measurements.forEach((measurement) => {
    const { block } = measurement;
    const lineEnds = measurement.paragraphLineEndOffsets;
    const lineHeights = measurement.paragraphLineHeights;
    if (
      block.type === 'paragraph' &&
      lineEnds &&
      lineHeights &&
      lineEnds.length === lineHeights.length &&
      lineEnds.length > 1
    ) {
      let lineStart = 0;
      while (lineStart < lineEnds.length) {
        if (blocks.length && height + lineHeights[lineStart] > capacity()) {
          push();
        }
        let lineEnd = lineStart;
        let fragmentHeight = 0;
        while (
          lineEnd < lineEnds.length &&
          (lineEnd === lineStart ||
            height + fragmentHeight + lineHeights[lineEnd] <= capacity())
        ) {
          fragmentHeight += lineHeights[lineEnd];
          lineEnd += 1;
        }
        const startOffset = lineStart ? lineEnds[lineStart - 1] : 0;
        blocks.push(
          sliceMeasuredParagraph(block, startOffset, lineEnds[lineEnd - 1]),
        );
        height += fragmentHeight;
        lineStart = lineEnd;
        if (lineStart < lineEnds.length) push();
      }
      return;
    }
    if (blocks.length && height + measurement.height > capacity()) push();
    blocks.push(block);
    height += measurement.height;
  });
  push();
  return {
    noteId: note.id,
    fragments: fragments.length
      ? fragments
      : [{ note, height: item.offsetHeight }],
  };
}

/** 从隐藏页面中的脚注节点生成可复用的分页片段。 */
export function measureDocxFootnotes(
  notes: readonly DocxNote[],
  article: HTMLElement | null | undefined,
  page: DocxPage,
) {
  return new Map(
    notes.flatMap((note) => {
      const item = article?.querySelector<HTMLElement>(
        `[data-office-word-note-id="footnote:${note.id}"]`,
      );
      return item
        ? [[note.id, createMeasuredFootnote(note, item, page)] as const]
        : [];
    }),
  );
}
