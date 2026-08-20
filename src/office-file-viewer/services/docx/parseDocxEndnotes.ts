import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import type { DocxParseContext } from './docxParsingContext';
import {
  parseDocxNotePart,
  type DocxNoteBlockReader,
} from './parseDocxFootnotes';

/** 读取 DOCX 尾注正文并过滤分隔符保留项。 */
export function parseDocxEndnotes(
  entries: OfficeEntryMap,
  context: DocxParseContext,
  readBlocks: DocxNoteBlockReader,
) {
  return parseDocxNotePart(entries, context, 'endnote', readBlocks);
}
