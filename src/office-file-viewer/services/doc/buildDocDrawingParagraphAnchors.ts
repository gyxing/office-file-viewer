import type { DocTextSegment } from './docParseTypes';
import type { DocTextStyle } from './types';

/** 绘图锚点所在段落的字符范围与页内纵向偏移。 */
export type DocDrawingParagraphAnchor = {
  /** 段落在正文字符流中的起始位置。 */
  charStart: number;
  /** 段落在正文字符流中的结束位置。 */
  charEnd: number;
  /** 段落锚点相对当前物理页顶部的估算距离。 */
  top: number;
};

/** 将 Word 行距统一换算为标准化渲染像素。 */
function resolveParagraphLineHeight(style: DocTextStyle | undefined) {
  const fontSize = style?.fontSize ?? 14;
  const lineHeight = style?.lineHeight;
  if (lineHeight === undefined) return fontSize;
  return lineHeight > 4 ? lineHeight : fontSize * lineHeight;
}

/** 估算只承载浮动形状锚点的段落推进高度。 */
function estimateAnchorParagraphHeight(style: DocTextStyle | undefined) {
  return (
    (style?.spacingBefore ?? 0) +
    resolveParagraphLineHeight(style) +
    (style?.spacingAfter ?? 0)
  );
}

/** 建立 SPA 段落相对页顶的偏移，供段落参考坐标转换使用。 */
export function buildDocDrawingParagraphAnchors(
  segments: readonly DocTextSegment[],
) {
  const anchors: DocDrawingParagraphAnchor[] = [];
  let charOffset = 0;
  let paragraphStart = 0;
  let paragraphTop = 0;
  let paragraphStyle: DocTextStyle | undefined;

  segments.forEach((segment) => {
    for (let index = 0; index < segment.text.length; index += 1) {
      paragraphStyle ??= segment.style;
      const character = segment.text[index];
      charOffset += 1;
      if (character === '\f') {
        paragraphStart = charOffset;
        paragraphTop = 0;
        paragraphStyle = undefined;
        continue;
      }
      if (character !== '\r') continue;
      anchors.push({
        charStart: paragraphStart,
        charEnd: charOffset,
        top: paragraphTop,
      });
      paragraphTop += estimateAnchorParagraphHeight(paragraphStyle);
      paragraphStart = charOffset;
      paragraphStyle = undefined;
    }
  });

  if (paragraphStart < charOffset) {
    anchors.push({
      charStart: paragraphStart,
      charEnd: charOffset,
      top: paragraphTop,
    });
  }
  return anchors;
}

/** 查找指定 SPA 字符位置所属段落的页内纵向偏移。 */
export function findDocDrawingParagraphTop(
  anchors: readonly DocDrawingParagraphAnchor[] | undefined,
  charOffset: number,
) {
  return (
    anchors?.find(
      (anchor) => charOffset >= anchor.charStart && charOffset < anchor.charEnd,
    )?.top ?? 0
  );
}
