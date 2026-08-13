import type { CSSProperties } from 'react';
import type { ReflectionStyle } from '../../../services/pptx/types';

function clamp01(value: number) {
  return Math.max(0, Math.min(1, value));
}

function resolveReflectionPlacement(reflection: ReflectionStyle) {
  const direction = (((reflection.direction ?? 90) % 360) + 360) % 360;
  return direction >= 45 && direction < 135
    ? 'below'
    : direction >= 135 && direction < 225
    ? 'left'
    : direction >= 225 && direction < 315
    ? 'above'
    : 'right';
}

function resolveReflectionMask(reflection: ReflectionStyle, placement: string) {
  const gradientDirection =
    placement === 'below'
      ? 'to bottom'
      : placement === 'above'
      ? 'to top'
      : placement === 'left'
      ? 'to left'
      : 'to right';
  const startOpacity = clamp01(reflection.startOpacity ?? 0.5);
  const endOpacity = clamp01(reflection.endOpacity ?? 0);
  const startPosition = clamp01(reflection.startPosition ?? 0);
  const endPosition = Math.max(
    startPosition,
    clamp01(reflection.endPosition ?? 1),
  );
  return `linear-gradient(${gradientDirection}, rgba(0, 0, 0, ${startOpacity}) ${
    startPosition * 100
  }%, rgba(0, 0, 0, ${endOpacity}) ${endPosition * 100}%, transparent 100%)`;
}

/** 将 OOXML 倒影参数转换为 Chromium 支持的盒倒影表达式。 */
export function reflectionToCss(reflection?: ReflectionStyle) {
  if (!reflection) return undefined;
  const placement = resolveReflectionPlacement(reflection);
  const gradientDirection =
    placement === 'below'
      ? 'to bottom'
      : placement === 'above'
      ? 'to top'
      : placement === 'left'
      ? 'to left'
      : 'to right';
  const startOpacity = clamp01(reflection.startOpacity ?? 0.5);
  const endOpacity = clamp01(reflection.endOpacity ?? 0);
  const startPosition = clamp01(reflection.startPosition ?? 0);
  const endPosition = Math.max(
    startPosition,
    clamp01(reflection.endPosition ?? 1),
  );
  return `${placement} ${Math.max(
    0,
    reflection.distance ?? 0,
  )}px linear-gradient(${gradientDirection}, rgba(255, 255, 255, ${startOpacity}) ${
    startPosition * 100
  }%, rgba(255, 255, 255, ${endOpacity}) ${endPosition * 100}%)`;
}

/** 生成不参与段落排版的文字倒影副本样式，避免盒倒影撑高行盒。 */
export function reflectionCopyToCss(
  reflection: ReflectionStyle,
): CSSProperties {
  const placement = resolveReflectionPlacement(reflection);
  const distance = Math.max(0, reflection.distance ?? 0);
  const maskImage = resolveReflectionMask(reflection, placement);
  const position: CSSProperties =
    placement === 'below'
      ? { left: 0, top: `calc(100% + ${distance}px)` }
      : placement === 'above'
      ? { left: 0, bottom: `calc(100% + ${distance}px)` }
      : placement === 'left'
      ? { right: `calc(100% + ${distance}px)`, top: 0 }
      : { left: `calc(100% + ${distance}px)`, top: 0 };
  return {
    position: 'absolute',
    ...position,
    width: '100%',
    height: '100%',
    transform:
      placement === 'below' || placement === 'above'
        ? 'scaleY(-1)'
        : 'scaleX(-1)',
    filter: reflection.blur ? `blur(${reflection.blur}px)` : undefined,
    WebkitMaskImage: maskImage,
    maskImage,
    pointerEvents: 'none',
    userSelect: 'none',
  };
}
