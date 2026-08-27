import type { DocxMeasuredBlock } from '../../services/docx/docxPagination';
import { canSplitMeasuredParagraph } from '../../services/docx/docxPagination';

/** 读取纯文本段落最后一个可见行内盒相对段落顶部的底边。 */
function measureVisibleInlineBottom(element: HTMLElement) {
  const elementTop = element.getBoundingClientRect().top;
  let bottom: number | undefined;
  element
    .querySelectorAll<HTMLElement>('span:not(.office-file-docx-auto-spacing)')
    .forEach((inline) => {
      Array.from(inline.getClientRects()).forEach((rect) => {
        if (rect.width <= 0 && rect.height <= 0) return;
        bottom = Math.max(bottom ?? rect.bottom, rect.bottom);
      });
    });
  return bottom === undefined ? undefined : Math.max(0, bottom - elementTop);
}

/** 读取纯文本段落的浏览器换行位置，供小文件物化分页精确拆行。 */
export function measureDocxParagraphLines(
  element: HTMLElement,
  block: DocxMeasuredBlock['block'],
  blockHeight: number,
): Pick<
  DocxMeasuredBlock,
  | 'pageEndHeight'
  | 'paragraphLineBoxEndHeight'
  | 'paragraphLineCount'
  | 'paragraphLineEndOffsets'
  | 'paragraphLineHeights'
> {
  if (block.type !== 'paragraph' || !canSplitMeasuredParagraph(block)) {
    return {};
  }
  const expectedText = block.inlines
    .map((inline) => (inline.type === 'text' ? inline.text : ''))
    .join('');
  if (!expectedText) return {};

  const walker = window.document.createTreeWalker(
    element,
    window.NodeFilter.SHOW_TEXT,
    {
      acceptNode(node) {
        return node.parentElement?.closest('.office-file-docx-auto-spacing')
          ? window.NodeFilter.FILTER_REJECT
          : window.NodeFilter.FILTER_ACCEPT;
      },
    },
  );
  const textNodes: Text[] = [];
  let currentTextNode = walker.nextNode() as Text | null;
  while (currentTextNode) {
    textNodes.push(currentTextNode);
    currentTextNode = walker.nextNode() as Text | null;
  }
  if (textNodes.map((node) => node.data).join('') !== expectedText) return {};

  const range = window.document.createRange();
  const lines: Array<{ top: number; bottom: number; endOffset: number }> = [];
  let textOffset = 0;
  for (const textNode of textNodes) {
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
  }
  range.detach();
  if (textOffset !== expectedText.length) return {};
  const elementTop = element.getBoundingClientRect().top;
  const visiblePageEndHeight =
    measureVisibleInlineBottom(element) ??
    (lines.length
      ? Math.max(0, lines[lines.length - 1].bottom - elementTop)
      : undefined);
  const lastLine = lines[lines.length - 1];
  const previousLine = lines[lines.length - 2];
  // Word 按行盒分页；Range 只返回字形盒，多行段尾需要补相邻行距中位于字形下方的一半留白。
  const trailingHalfLeading =
    lastLine && previousLine
      ? Math.max(
          0,
          (lastLine.top - previousLine.top - (lastLine.bottom - lastLine.top)) /
            2,
        )
      : 0;
  const paragraphLineBoxEndHeight =
    visiblePageEndHeight === undefined
      ? undefined
      : visiblePageEndHeight + trailingHalfLeading;
  const pageEndHeight = visiblePageEndHeight;
  if (lines.length < 2) {
    return { pageEndHeight, paragraphLineCount: lines.length };
  }
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

  // 拆页后的单行段落仍占完整行盒；按相邻行顶部距离计量，避免中点算法把首行压成半行。
  const lineHeights = lines.map((line, index) =>
    index < lines.length - 1 ? Math.max(0, lines[index + 1].top - line.top) : 0,
  );
  lineHeights[lineHeights.length - 1] =
    blockHeight -
    lineHeights.slice(0, -1).reduce((sum, height) => sum + height, 0);
  if (lineHeights.some((height) => height <= 0)) return {};
  return {
    pageEndHeight,
    paragraphLineBoxEndHeight,
    paragraphLineCount: lines.length,
    paragraphLineEndOffsets: lineEndOffsets,
    paragraphLineHeights: lineHeights,
  };
}
