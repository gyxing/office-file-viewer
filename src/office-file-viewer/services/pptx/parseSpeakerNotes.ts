import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import { readXml } from '../../shared/ooxml/archive';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  descendantByLocalName,
  descendantsByLocalName,
  matchesLocalName,
  parseXml,
  textContent,
} from '../../shared/ooxml/xml';
import type {
  SpeakerNotesModel,
  TextParagraph,
  TextRun,
} from '../presentation/types';
import type { RelationshipMap } from './PptxPackageContext';

/** 将备注页中的一个正文段落转换为统一文本模型。 */
function parseNotesParagraph(node: Element): TextParagraph | undefined {
  const paragraphProperties = childByLocalName(node, 'pPr');
  const bulletChar = attr(
    childByLocalName(paragraphProperties, 'buChar'),
    'char',
  );
  const bulletFontFamily = attr(
    childByLocalName(paragraphProperties, 'buFont'),
    'typeface',
  );
  const bulletNone = Boolean(childByLocalName(paragraphProperties, 'buNone'));
  const runs: TextRun[] = [];

  Array.from(node.children).forEach((child) => {
    if (matchesLocalName(child, 'r') || matchesLocalName(child, 'fld')) {
      runs.push({ text: textContent(childByLocalName(child, 't')) });
    } else if (matchesLocalName(child, 'br')) {
      runs.push({ text: '\n' });
    }
  });

  if (!runs.length && textContent(node)) runs.push({ text: textContent(node) });
  if (!runs.some((run) => run.text.length > 0)) return undefined;

  return {
    level: Number(attr(paragraphProperties, 'lvl') ?? 0),
    runs,
    bullet:
      bulletChar || bulletNone
        ? {
            char: bulletChar,
            fontFamily: bulletFontFamily,
            none: bulletNone || undefined,
          }
        : undefined,
  };
}

/** 从 PPTX 当前幻灯片关系中读取演讲者正文，过滤日期、页脚等备注页占位符。 */
export function parsePptxSpeakerNotes(
  entries: OfficeEntryMap,
  relationships: RelationshipMap,
  slideRelsPath: string,
): SpeakerNotesModel | undefined {
  const notesPath = Object.values(relationships[slideRelsPath] ?? {}).find(
    (relationship) => relationship.target.includes('notesSlides/'),
  )?.target;
  if (!notesPath) return undefined;

  const xml = readXml(entries, notesPath);
  if (!xml) return undefined;
  const document = parseXml(xml);
  const paragraphs = descendantsByLocalName(document.documentElement, 'sp')
    .filter(
      (shape) => attr(descendantByLocalName(shape, 'ph'), 'type') === 'body',
    )
    .flatMap((shape) =>
      childrenByLocalName(descendantByLocalName(shape, 'txBody'), 'p'),
    )
    .map(parseNotesParagraph)
    .filter((paragraph): paragraph is TextParagraph => Boolean(paragraph));
  const plainText = paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n');

  return plainText.trim() ? { paragraphs, plainText } : undefined;
}
