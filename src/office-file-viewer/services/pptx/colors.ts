import type { ThemeModel } from './types';

/** 将输入标准化为 `normalizeColor` 返回的结构。 */
export function normalizeColor(input?: string) {
  if (!input) {
    return undefined;
  }

  return input.trim();
}

/** 将来源颜色规范化为六位十六进制颜色。 */
export function toHexColor(input?: string) {
  const value = normalizeColor(input);
  if (!value) {
    return undefined;
  }

  if (value.startsWith('#')) {
    return value;
  }

  if (/^[0-9a-f]{6}$/i.test(value)) {
    return `#${value}`;
  }

  return value;
}

/** 解析并确定 `resolveThemeColor` 对应的引用或配置。 */
export function resolveThemeColor(name: string | undefined, theme: ThemeModel) {
  if (!name) {
    return undefined;
  }

  const mapped = theme.colorMap?.[name] ?? name;
  return toHexColor(
    theme.colorScheme[mapped] ?? theme.colorScheme[name] ?? mapped,
  );
}

function clamp255(value: number) {
  return Math.max(0, Math.min(255, value));
}

function hexToRgb(hex: string) {
  const normalized = hex.replace('#', '');
  if (!/^[0-9a-f]{6}$/i.test(normalized)) {
    return null;
  }

  const value = Number.parseInt(normalized, 16);
  return {
    r: (value >> 16) & 255,
    g: (value >> 8) & 255,
    b: value & 255,
  };
}

function rgbToHex(r: number, g: number, b: number) {
  return `#${[r, g, b]
    .map((value) => clamp255(Math.round(value)).toString(16).padStart(2, '0'))
    .join('')}`;
}

/** 按 OOXML 颜色变换规则计算最终颜色。 */
export function transformColor(
  hex: string | undefined,
  transforms: Array<{
    /** 用于区分联合类型分支的类型标识。 */
    type: string;
    /** OOXML 颜色变换的参数值。 */
    val: number;
  }>,
) {
  if (!hex) {
    return undefined;
  }

  const rgb = hexToRgb(hex);
  if (!rgb) {
    return hex;
  }

  let { r, g, b } = rgb;
  transforms.forEach((transform) => {
    if (transform.type === 'tint') {
      const ratio = transform.val / 100000;
      r = r + (255 - r) * ratio;
      g = g + (255 - g) * ratio;
      b = b + (255 - b) * ratio;
    }
    if (transform.type === 'shade') {
      const ratio = transform.val / 100000;
      r = r * ratio;
      g = g * ratio;
      b = b * ratio;
    }
    if (transform.type === 'lumMod') {
      const ratio = transform.val / 100000;
      r *= ratio;
      g *= ratio;
      b *= ratio;
    }
    if (transform.type === 'lumOff') {
      const ratio = transform.val / 100000;
      r += 255 * ratio;
      g += 255 * ratio;
      b += 255 * ratio;
    }
  });

  return rgbToHex(r, g, b);
}

/** 将 OOXML 透明度值转换为零到一的透明度。 */
export function alphaToOpacity(alpha?: string) {
  if (!alpha) {
    return undefined;
  }

  const value = Number(alpha);
  if (!Number.isFinite(value)) {
    return undefined;
  }

  return Math.max(0, Math.min(1, value / 100000));
}

/** 将 OOXML 透明度值转换为零到一的比例。 */
export function alphaToRatio(alpha?: string) {
  if (!alpha) {
    return undefined;
  }

  const value = Number(alpha);
  if (!Number.isFinite(value)) {
    return undefined;
  }

  return Math.max(0, Math.min(1, value / 100000));
}
