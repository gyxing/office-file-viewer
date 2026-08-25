import type {
  OfficeViewerWatermark,
  ResolvedOfficeViewerWatermark,
} from './types';

/** 约束数值属性，避免异常传参产生不可用 SVG 尺寸。 */
function clampNumber(
  value: number | undefined,
  fallback: number,
  minimum: number,
  maximum: number,
): number {
  if (value === undefined || !Number.isFinite(value)) return fallback;
  return Math.min(maximum, Math.max(minimum, value));
}

/** 读取有限数值，偏移量允许使用负值。 */
function finiteNumber(value: number | undefined, fallback: number): number {
  return value !== undefined && Number.isFinite(value) ? value : fallback;
}

/** 清理水印文本并为全部渲染属性提供稳定默认值。 */
export function resolveOfficeViewerWatermark(
  watermark: OfficeViewerWatermark | undefined,
): ResolvedOfficeViewerWatermark | undefined {
  if (!watermark) return undefined;
  const sourceContent = Array.isArray(watermark.content)
    ? watermark.content
    : [watermark.content];
  const content = sourceContent
    .map((line) => String(line).trim())
    .filter(Boolean);
  if (!content.length) return undefined;

  return {
    content,
    color: watermark.color ?? '#64748b',
    opacity: clampNumber(watermark.opacity, 0.12, 0, 1),
    fontSize: clampNumber(watermark.fontSize, 16, 8, 96),
    fontFamily: watermark.fontFamily ?? 'sans-serif',
    fontWeight: watermark.fontWeight ?? 400,
    rotate: finiteNumber(watermark.rotate, -22),
    gap: [
      clampNumber(watermark.gap?.[0], 120, 0, 1200),
      clampNumber(watermark.gap?.[1], 96, 0, 1200),
    ],
    offset: [
      finiteNumber(watermark.offset?.[0], 0),
      finiteNumber(watermark.offset?.[1], 0),
    ],
  };
}
