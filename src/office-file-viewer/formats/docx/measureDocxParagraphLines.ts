import type { DocxMeasuredBlock } from '../../services/docx/docxPagination';
import { canSplitMeasuredParagraph } from '../../services/docx/docxPagination';

/** 读取纯文本段落的浏览器换行位置，供小文件物化分页精确拆行。 */
export function measureDocxParagraphLines(
  element: HTMLElement,
  block: DocxMeasuredBlock['block'],
  blockHeight: number,
): Pick<DocxMeasuredBlock, 'paragraphLineEndOffsets' | 'paragraphLineHeights'> {
  if (block.type !== 'paragraph' || !canSplitMeasuredParagraph(block)) {
    return {};
  }
  const expectedText = block.inlines
    .map((inline) => (inline.type === 'text' ? inline.text : ''))
    .join('');
  if (!expectedText || element.textContent !== expectedText) return {};

  const range = window.document.createRange();
  const walker = window.document.createTreeWalker(
    element,
    window.NodeFilter.SHOW_TEXT,
  );
  const lines: Array<{ top: number; bottom: number; endOffset: number }> = [];
  let textOffset = 0;
  let textNode = walker.nextNode() as Text | null;
  while (textNode) {
    const textLength = textNode.data.length;
    for (
      let characterOffset = 1;
      characterOffset <= textLength;
      characterOffset += 1
    ) {
      range.setStart(textNode, characterOffset - 1);
      range.setEnd(textNode, characterOffset);
      const rect = Array.from(range.getClientRects()).find(
        (item) => item.width > 0 || item.height > 0,
      );
      const endOffset = textOffset + characterOffset;
      if (!rect) {
        if (lines.length) lines[lines.length - 1].endOffset = endOffset;
        continue;
      }
      const currentLine = lines[lines.length - 1];
      const verticalOverlap = currentLine
        ? Math.min(currentLine.bottom, rect.bottom) -
          Math.max(currentLine.top, rect.top)
        : 0;
      if (currentLine && verticalOverlap <= 0.5 && rect.top > currentLine.top) {
        currentLine.endOffset = Math.max(currentLine.endOffset, endOffset - 1);
        lines.push({ top: rect.top, bottom: rect.bottom, endOffset });
      } else if (currentLine) {
        currentLine.top = Math.min(currentLine.top, rect.top);
        currentLine.bottom = Math.max(currentLine.bottom, rect.bottom);
        currentLine.endOffset = endOffset;
      } else {
        lines.push({ top: rect.top, bottom: rect.bottom, endOffset });
      }
    }
    textOffset += textLength;
    textNode = walker.nextNode() as Text | null;
  }
  range.detach();
  if (textOffset !== expectedText.length || lines.length < 2) return {};
  lines[lines.length - 1].endOffset = textOffset;
  const lineEndOffsets = lines
    .map((line) => line.endOffset)
    .filter((offset, index, offsets) => offset > (offsets[index - 1] ?? 0));
  if (
    lineEndOffsets.length !== lines.length ||
    lineEndOffsets[lineEndOffsets.length - 1] !== expectedText.length
  ) {
    return {};
  }

  const elementRect = element.getBoundingClientRect();
  const lineBoundaries = [
    elementRect.top,
    ...lines
      .slice(0, -1)
      .map((line, index) => (line.top + lines[index + 1].top) / 2),
    elementRect.bottom,
  ];
  const lineHeights = lines.map((_, index) =>
    Math.max(0, lineBoundaries[index + 1] - lineBoundaries[index]),
  );
  lineHeights[lineHeights.length - 1] +=
    blockHeight - lineHeights.reduce((sum, height) => sum + height, 0);
  if (lineHeights.some((height) => height <= 0)) return {};
  return {
    paragraphLineEndOffsets: lineEndOffsets,
    paragraphLineHeights: lineHeights,
  };
}
