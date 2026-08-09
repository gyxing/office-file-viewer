import type {
  TextParagraph,
  TextRun,
  TextStyle,
} from '../../presentation/types';
import {
  readPptTextHyperlinkRanges,
  type PptTextHyperlinkRange,
} from '../document/readHyperlinks';
import type { PptParseContext, PptRecord } from '../types';
import { mergePptTextStyles } from './mergeTextStyles';
import { readPptTextAtoms } from './readTextAtoms';
import { readPptTextStyles } from './readTextStyles';
import type {
  PptCharacterStyleRun,
  PptParsedText,
  PptTextDefaults,
} from './types';

function styleAt(
  runs: PptCharacterStyleRun[],
  position: number,
): TextStyle | undefined {
  let end = 0;
  for (const run of runs) {
    end += run.count;
    if (position < end) return run.style;
  }
  return runs.length ? runs[runs.length - 1].style : undefined;
}

function paragraphStyleAt(
  runs: ReturnType<typeof readPptTextStyles>['paragraphs'],
  position: number,
) {
  let end = 0;
  for (const run of runs) {
    end += run.count;
    if (position < end) return run;
  }
  return runs.length ? runs[runs.length - 1] : undefined;
}

function hyperlinkAt(
  ranges: readonly PptTextHyperlinkRange[],
  position: number,
) {
  return ranges.find((range) => position >= range.begin && position < range.end)
    ?.hyperlink;
}

function appendRuns(
  text: string,
  start: number,
  characterRuns: PptCharacterStyleRun[],
  baseStyle: TextStyle,
  hyperlinkRanges: readonly PptTextHyperlinkRange[],
) {
  const runs: TextRun[] = [];
  let cursor = 0;
  while (cursor < text.length) {
    const style = styleAt(characterRuns, start + cursor);
    const hyperlink = hyperlinkAt(hyperlinkRanges, start + cursor);
    let end = cursor + 1;
    while (
      end < text.length &&
      styleAt(characterRuns, start + end) === style &&
      hyperlinkAt(hyperlinkRanges, start + end) === hyperlink
    ) {
      end += 1;
    }
    runs.push({
      text: text.slice(cursor, end),
      style: mergePptTextStyles(baseStyle, style),
      hyperlink,
    });
    cursor = end;
  }
  if (!runs.length) runs.push({ text: '', style: baseStyle });
  return runs;
}

function buildParagraphs(
  text: string,
  defaults: PptTextDefaults,
  styles: ReturnType<typeof readPptTextStyles>,
  hyperlinkRanges: readonly PptTextHyperlinkRange[],
) {
  const paragraphs: TextParagraph[] = [];
  const baseStyle = mergePptTextStyles(
    defaults.document,
    defaults.master,
    defaults.placeholder,
  );
  let position = 0;
  for (const value of text.replace(/\u000b/g, '\n').split(/\r\n?|\n/)) {
    const paragraphRun = paragraphStyleAt(styles.paragraphs, position);
    const paragraphStyle = mergePptTextStyles(baseStyle, paragraphRun?.style);
    paragraphs.push({
      runs: appendRuns(
        value,
        position,
        styles.characters,
        paragraphStyle,
        hyperlinkRanges,
      ),
      style: paragraphStyle,
      level: paragraphRun?.level ?? 0,
      bullet: paragraphStyle.bullet,
    });
    position += value.length + 1;
  }
  return paragraphs;
}

/** 将文本框子记录解析为统一演示文稿文本段落。 */
export function parsePptTextGroups(
  records: PptRecord[],
  defaults: PptTextDefaults,
  context: PptParseContext,
): PptParsedText[] {
  const hyperlinkRanges = readPptTextHyperlinkRanges(records, context);
  return readPptTextAtoms(records, context).map((group, groupIndex) => ({
    textType: group.textType,
    paragraphs: buildParagraphs(
      group.text,
      defaults,
      readPptTextStyles(
        group.styleRecord,
        group.text.length,
        defaults,
        context,
      ),
      hyperlinkRanges.filter((range) => range.groupIndex === groupIndex),
    ),
  }));
}

/** 兼容单文本框调用约定，返回其中所有段落。 */
export function parsePptText(
  records: PptRecord[],
  defaults: PptTextDefaults,
  context: PptParseContext,
): TextParagraph[] {
  return parsePptTextGroups(records, defaults, context).flatMap(
    (group) => group.paragraphs,
  );
}
