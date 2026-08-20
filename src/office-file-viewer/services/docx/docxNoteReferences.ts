import type {
  DocxBlock,
  DocxNote,
  DocxNoteReferenceInline,
  DocxPageContent,
} from './types';

/** 递归收集正文块中的脚注和尾注引用。 */
export function collectDocxNoteReferences(blocks: readonly DocxBlock[]) {
  const references: DocxNoteReferenceInline[] = [];
  const visitBlock = (block: DocxBlock) => {
    if (block.type === 'paragraph') {
      block.inlines.forEach((inline) => {
        if (inline.type === 'note-reference') references.push(inline);
        if (inline.type === 'shape') {
          inline.shape.items.forEach((item) =>
            (item.blocks ?? item.paragraphs ?? []).forEach(visitBlock),
          );
        }
      });
      return;
    }
    if (block.type === 'table') {
      block.rows.forEach((row) =>
        row.cells.forEach((cell) => cell.blocks.forEach(visitBlock)),
      );
    }
  };
  blocks.forEach(visitBlock);
  return references;
}

/** 为单个源页面附加实际被正文引用的去重脚注。 */
export function attachDocxFootnotesToPage(
  page: DocxPageContent,
  footnotes: readonly DocxNote[],
): DocxPageContent {
  if (!footnotes.length) return page;
  const requestedIds = new Set(
    collectDocxNoteReferences(page.blocks)
      .filter((reference) => reference.noteKind === 'footnote')
      .map((reference) => reference.noteId),
  );
  if (!requestedIds.size) return page;
  const attached = footnotes.filter((note) => requestedIds.has(note.id));
  return attached.length ? { ...page, footnotes: attached } : page;
}
