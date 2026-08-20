import type { CSSProperties } from 'react';
import React from 'react';
import type { DocxNote, DocxPage } from '../../services/docx/types';
import { DocxBlockRenderer } from '../docx/DocxBlockRenderer';
export { collectDocxNoteReferences } from '../../services/docx/docxNoteReferences';

/** 页面脚注或文档尾注渲染属性。 */
type WordNoteBlockProps = {
  /** 当前页面需要显示的脚注或尾注。 */
  notes: readonly DocxNote[];
  /** 当前页面尺寸和页边距。 */
  page: DocxPage;
  /** 尾注使用普通文档末尾布局，脚注贴近当前页底部。 */
  endnotes?: boolean;
};

/** 在页面底部或文档末尾渲染脚注、尾注正文。 */
export function WordNoteBlock({
  notes,
  page,
  endnotes = false,
}: WordNoteBlockProps) {
  if (!notes.length) return null;
  const contentWidth = page.width - page.marginLeft - page.marginRight;
  const style = endnotes
    ? undefined
    : ({
        right: page.marginRight,
        bottom: page.marginBottom,
        left: page.marginLeft,
      } as CSSProperties);
  return (
    <section
      className={
        endnotes
          ? 'office-file-word-notes office-file-word-notes--endnotes'
          : 'office-file-word-notes office-file-word-notes--footnotes'
      }
      style={style}
    >
      <span className="office-file-word-notes__separator" aria-hidden="true" />
      {notes.map((note) => (
        <div
          key={`${note.kind}-${note.id}`}
          className="office-file-word-notes__item"
          data-office-word-note-id={`${note.kind}:${note.id}`}
        >
          <sup>{note.label}</sup>
          <div className="office-file-word-notes__content">
            {note.blocks.map((block) => (
              <DocxBlockRenderer
                key={block.id}
                block={block}
                availableWidth={contentWidth}
                maximumWidth={page.width}
                suppressSpacingBefore
                suppressSpacingAfter
              />
            ))}
          </div>
        </div>
      ))}
    </section>
  );
}
