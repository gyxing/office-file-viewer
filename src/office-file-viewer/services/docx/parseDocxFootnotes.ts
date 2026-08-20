import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import { readXml } from '../../shared/ooxml/archive';
import { attr, descendantsByLocalName, parseXml } from '../../shared/ooxml/xml';
import type { DocxParseContext } from './docxParsingContext';
import type { DocxBlock, DocxNote } from './types';

/** 注释正文复用 DOCX 块解析器时使用的回调。 */
export type DocxNoteBlockReader = (
  container: Element,
  idPrefix: string,
  context: DocxParseContext,
) => DocxBlock[];

/** 建立源注释标识到页面连续编号的映射。 */
export function readDocxNoteLabelMap(
  entries: OfficeEntryMap,
  kind: DocxNote['kind'],
) {
  const partName = kind === 'footnote' ? 'footnotes' : 'endnotes';
  const elementName = kind === 'footnote' ? 'footnote' : 'endnote';
  const xml = readXml(entries, `word/${partName}.xml`);
  if (!xml) return {};
  const elements = descendantsByLocalName(
    parseXml(xml).documentElement,
    elementName,
  ).filter((element) => {
    const type = attr(element, 'w:type') ?? attr(element, 'type');
    const id = attr(element, 'w:id') ?? attr(element, 'id') ?? '';
    return !type && !id.startsWith('-');
  });
  return Object.fromEntries(
    elements.map((element, index) => [
      attr(element, 'w:id') ?? attr(element, 'id') ?? `${index + 1}`,
      `${index + 1}`,
    ]),
  );
}

/** 解析脚注和尾注部件共用的结构读取流程。 */
export function parseDocxNotePart(
  entries: OfficeEntryMap,
  context: DocxParseContext,
  kind: DocxNote['kind'],
  readBlocks: DocxNoteBlockReader,
): DocxNote[] {
  const partName = kind === 'footnote' ? 'footnotes' : 'endnotes';
  const xml = readXml(entries, `word/${partName}.xml`);
  if (!xml) return [];
  const relationships =
    context.packageState.relationships[`word/_rels/${partName}.xml.rels`] ?? {};
  const noteContext: DocxParseContext = {
    ...context,
    documentRels: relationships,
  };
  const elementName = kind === 'footnote' ? 'footnote' : 'endnote';
  const noteElements = descendantsByLocalName(
    parseXml(xml).documentElement,
    elementName,
  ).filter((element) => {
    const type = attr(element, 'w:type') ?? attr(element, 'type');
    const id = attr(element, 'w:id') ?? attr(element, 'id') ?? '';
    return !type && !id.startsWith('-');
  });

  return noteElements.map((element, index) => {
    const id = attr(element, 'w:id') ?? attr(element, 'id') ?? `${index + 1}`;
    return {
      id,
      kind,
      label: `${index + 1}`,
      blocks: readBlocks(element, `docx-${kind}-${id}`, noteContext),
    };
  });
}

/** 读取 DOCX 脚注正文并过滤分隔符保留项。 */
export function parseDocxFootnotes(
  entries: OfficeEntryMap,
  context: DocxParseContext,
  readBlocks: DocxNoteBlockReader,
) {
  return parseDocxNotePart(entries, context, 'footnote', readBlocks);
}
