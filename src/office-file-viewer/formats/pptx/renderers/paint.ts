// paint 工具负责把 PPTX 颜色、透明度和渐变模型转换成 CSS/SVG 可用格式。
import type { GradientFill } from '../../../services/pptx/types';

/** 将独立透明度合并到颜色值。 */
export function colorWithOpacity(color?: string, opacity?: number) {
  if (!color || opacity === undefined || opacity >= 1) return color;
  const normalized = color.replace('#', '');
  if (!/^[0-9a-f]{6}$/i.test(normalized)) return color;
  const value = Number.parseInt(normalized, 16);
  const r = (value >> 16) & 255;
  const g = (value >> 8) & 255;
  const b = value & 255;
  return `rgba(${r}, ${g}, ${b}, ${opacity})`;
}

/** 判断当前填充是否为渐变。 */
export function isGradientPaint(
  paint?: string | GradientFill | null,
): paint is GradientFill {
  return Boolean(paint && typeof paint === 'object' && paint.type === 'linear');
}

/** 将输入标准化为 `normalizeCssAngle` 返回的结构。 */
function normalizeCssAngle(angle: number) {
  return (((angle + 90) % 360) + 360) % 360;
}

/** 把输入格式化为 `formatOffset` 返回的展示值。 */
function formatOffset(offset: number) {
  return `${Math.max(0, Math.min(100, offset * 100))
    .toFixed(1)
    .replace(/\.0$/, '')}%`;
}

/** 将演示文稿填充转换为 CSS 颜色或渐变。 */
export function paintToCss(
  paint?: string | GradientFill | null,
  opacity?: number,
) {
  if (!paint) return undefined;
  if (!isGradientPaint(paint)) return colorWithOpacity(paint, opacity);
  const stops = paint.stops
    .slice()
    .sort((a, b) => a.offset - b.offset)
    .map((stop) => `${stop.color} ${formatOffset(stop.offset)}`);
  return `linear-gradient(${normalizeCssAngle(paint.angle)}deg, ${stops.join(
    ', ',
  )})`;
}

/** 将渐变角度换算为 SVG 起止坐标。 */
export function gradientToSvgEndpoints(angle: number) {
  const radians = (angle * Math.PI) / 180;
  return {
    x1: 0.5 - Math.cos(radians) / 2,
    y1: 0.5 - Math.sin(radians) / 2,
    x2: 0.5 + Math.cos(radians) / 2,
    y2: 0.5 + Math.sin(radians) / 2,
  };
}
/** 将 DrawingML 预设线型转换为 CSS 边框线型。 */
export function presentationLineStyle(dash?: string) {
  const normalized = dash?.toLowerCase();
  if (!normalized || normalized === 'solid') return 'solid';
  return normalized.includes('dot') && !normalized.includes('dash')
    ? 'dotted'
    : 'dashed';
}

/** 将 DrawingML 预设线型转换为 SVG 可识别的虚线数组。 */
export function presentationStrokeDasharray(dash?: string, strokeWidth = 1) {
  const normalized = dash?.toLowerCase();
  if (!normalized || normalized === 'solid') return undefined;

  const width = Math.max(1, strokeWidth);
  const units = (...values: number[]) =>
    values.map((value) => value * width).join(' ');
  switch (normalized) {
    case 'dot':
    case 'sysdot':
      return units(0, 2);
    case 'sysdash':
      return units(3, 2);
    case 'dashdot':
    case 'sysdashdot':
      return units(4, 3, 0, 3);
    case 'lgdash':
      return units(8, 3);
    case 'lgdashdot':
      return units(8, 3, 0, 3);
    case 'lgdashdotdot':
    case 'sysdashdotdot':
      return units(8, 3, 0, 3, 0, 3);
    default:
      return units(4, 3);
  }
}
