import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import { readXml } from '../../shared/ooxml/archive';
import type { OfficeRelationship } from '../../shared/ooxml/media';
import {
  attr,
  childrenByLocalName,
  descendantsByLocalName,
  parseXml,
  textContent,
} from '../../shared/ooxml/xml';
import type { SpreadsheetAnnotation } from '../spreadsheet/semantics/types';
import type {
  XlsxPackageContext,
  XlsxSheetDescriptor,
} from './XlsxPackageContext';
import { decodeMojibake, parseCellRef } from './xlsxCellFormatting';

/** 判断关系是否指向经典或线程化批注部件。 */
function commentRelationshipKind(relationship: OfficeRelationship) {
  if (/threadedComment/i.test(relationship.type ?? relationship.target)) {
    return 'threaded' as const;
  }
  if (/\/comments$/i.test(relationship.type ?? '')) return 'classic' as const;
  if (/^xl\/comments\d*\.xml$/i.test(relationship.target)) {
    return 'classic' as const;
  }
  return undefined;
}

/** 读取批注中的富文本并还原潜在乱码。 */
function readCommentText(node: Element) {
  return decodeMojibake(
    descendantsByLocalName(node, 't')
      .map((text) => textContent(text))
      .join(''),
  );
}

/** 解析经典 comments.xml。 */
function parseClassicComments(
  xml: string,
  sheetId: string,
): SpreadsheetAnnotation[] {
  if (!xml) return [];
  const root = parseXml(xml).documentElement;
  const authors = childrenByLocalName(
    childrenByLocalName(root, 'authors')[0],
    'author',
  ).map((author) => decodeMojibake(textContent(author)));
  return descendantsByLocalName(root, 'comment').flatMap((comment, index) => {
    const ref = attr(comment, 'ref');
    if (!ref) return [];
    const address = parseCellRef(ref);
    const authorId = Number(attr(comment, 'authorId'));
    return [
      {
        id: `${sheetId}:comment:${index + 1}`,
        ref,
        row: address.row,
        column: address.column,
        author: Number.isInteger(authorId) ? authors[authorId] : undefined,
        text: readCommentText(comment),
      },
    ];
  });
}

/** 解析 persons/person.xml 中的线程批注作者。 */
function parsePersons(xml: string) {
  if (!xml) return new Map<string, string>();
  return new Map(
    descendantsByLocalName(parseXml(xml).documentElement, 'person').flatMap(
      (person) => {
        const id = attr(person, 'id');
        const name = attr(person, 'displayName') ?? attr(person, 'userId');
        return id && name ? [[id, decodeMojibake(name)] as const] : [];
      },
    ),
  );
}

/** 解析 threadedComments.xml 并保留回复关系。 */
function parseThreadedComments(
  xml: string,
  sheetId: string,
  persons: ReadonlyMap<string, string>,
): SpreadsheetAnnotation[] {
  if (!xml) return [];
  return descendantsByLocalName(
    parseXml(xml).documentElement,
    'threadedComment',
  ).flatMap((comment, index) => {
    const ref = attr(comment, 'ref');
    if (!ref) return [];
    const address = parseCellRef(ref);
    const personId = attr(comment, 'personId');
    return [
      {
        id: attr(comment, 'id') ?? `${sheetId}:threaded-comment:${index + 1}`,
        ref,
        row: address.row,
        column: address.column,
        author: personId ? persons.get(personId) : undefined,
        createdAt: attr(comment, 'dT'),
        parentId: attr(comment, 'parentId'),
        resolved: attr(comment, 'done') === '1',
        text: readCommentText(comment),
      },
    ];
  });
}

/** 找出物化包中的线程批注作者部件。 */
function findMaterializedPersons(entries: OfficeEntryMap) {
  const path = [...entries.keys()].find((entry) =>
    /^xl\/persons\/person.*\.xml$/i.test(entry),
  );
  return path ? readXml(entries, path) : '';
}

/** 读取物化工作表关系指向的全部批注。 */
export function parseMaterializedXlsxComments(input: {
  /** 当前工作表标识。 */
  sheetId: string;
  /** 当前工作表关系。 */
  relationships: Record<string, OfficeRelationship>;
  /** 当前 OOXML 包条目。 */
  entries: OfficeEntryMap;
}) {
  const persons = parsePersons(findMaterializedPersons(input.entries));
  return Object.values(input.relationships).flatMap((relationship) => {
    const kind = commentRelationshipKind(relationship);
    if (!kind) return [];
    const xml = readXml(input.entries, relationship.target);
    return kind === 'classic'
      ? parseClassicComments(xml, input.sheetId)
      : parseThreadedComments(xml, input.sheetId, persons);
  });
}

/** 按需读取当前工作表关系指向的全部批注。 */
export async function parseSourceXlsxComments(
  context: XlsxPackageContext,
  descriptor: XlsxSheetDescriptor,
  signal?: AbortSignal,
) {
  const relationships = Object.values(
    context.relationships[descriptor.relsPath] ?? {},
  );
  const personEntry = context.reader
    .list('xl/persons/')
    .find((entry) => /^xl\/persons\/person.*\.xml$/i.test(entry.path));
  const persons = parsePersons(
    personEntry ? await context.reader.readText(personEntry.path, signal) : '',
  );
  const annotations: SpreadsheetAnnotation[] = [];
  for (const relationship of relationships) {
    const kind = commentRelationshipKind(relationship);
    if (!kind) continue;
    const xml = await context.reader.readText(relationship.target, signal);
    annotations.push(
      ...(kind === 'classic'
        ? parseClassicComments(xml, descriptor.id)
        : parseThreadedComments(xml, descriptor.id, persons)),
    );
  }
  return annotations;
}
