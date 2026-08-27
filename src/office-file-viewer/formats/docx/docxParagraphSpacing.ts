import type { DocxBlock } from '../../services/docx/types';

/** 判断相邻段落是否应按 OOXML contextualSpacing 取消两者之间的间距。 */
export function shouldSuppressDocxContextualSpacing(
  block: DocxBlock | undefined,
  adjacent: DocxBlock | undefined,
) {
  return Boolean(
    block?.type === 'paragraph' &&
      adjacent?.type === 'paragraph' &&
      block.contextualSpacing &&
      block.paragraphStyleId &&
      block.paragraphStyleId === adjacent.paragraphStyleId,
  );
}

/** 字体度量切换时补回被 CSS 外边距折叠吞掉的段前距。 */
export function resolveDocxSpacingBefore(
  block: DocxBlock,
  previousBlock: DocxBlock | undefined,
  suppressSpacingBefore: boolean,
) {
  if (block.type !== 'paragraph' || suppressSpacingBefore) return 0;
  const spacingBefore = block.spacingBefore ?? 0;
  if (
    spacingBefore <= 0 ||
    previousBlock?.type !== 'paragraph' ||
    shouldSuppressDocxContextualSpacing(previousBlock, block)
  ) {
    return spacingBefore;
  }
  const currentFontSize = block.style?.fontSize ?? 0;
  const previousFontSize = previousBlock.style?.fontSize ?? 0;
  if (
    block.paragraphStyleId &&
    block.paragraphStyleId === previousBlock.paragraphStyleId
  ) {
    return spacingBefore + (previousBlock.spacingAfter ?? 0);
  }
  // 同字号正文沿用 Word 的段落间距折叠；字体度量切换时才补回被浏览器吞掉的段前距。
  if (Math.abs(currentFontSize - previousFontSize) < 0.01) {
    return spacingBefore;
  }
  return spacingBefore + (previousBlock.spacingAfter ?? 0);
}
