import type { CSSProperties } from 'react';

/** 将演示文稿行高转换为 React 可正确识别的 CSS 值。 */
export function resolvePresentationLineHeight(
  lineHeight: number | undefined,
  fallback = 1.2,
): CSSProperties['lineHeight'] {
  const value = lineHeight ?? fallback;

  // OOXML 固定磅值行距会转换成大于 4 的像素值；数字直接传给 React 会被误当成倍数。
  return value > 4 ? `${value}px` : value;
}

/** 补偿 PowerPoint 与 CSS 对大行距首行上方留白分配方式的差异。 */
export function resolvePresentationLeadingOffset(
  fontSize: number | undefined,
  lineHeight: CSSProperties['lineHeight'],
) {
  if (!fontSize || lineHeight === undefined || lineHeight === 'normal') {
    return undefined;
  }
  const lineHeightPx =
    typeof lineHeight === 'number'
      ? lineHeight * fontSize
      : Number.parseFloat(String(lineHeight));
  if (!Number.isFinite(lineHeightPx) || lineHeightPx <= fontSize * 1.5) {
    return undefined;
  }

  return (lineHeightPx - fontSize) * 0.3;
}
