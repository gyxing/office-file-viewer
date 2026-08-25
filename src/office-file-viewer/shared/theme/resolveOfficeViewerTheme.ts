import type {
  OfficeViewerThemeMode,
  OfficeViewerThemeOptions,
  OfficeViewerThemeStyle,
} from './types';

/** 可由主题令牌直接映射的 CSS 变量。 */
const THEME_TOKEN_VARIABLES = {
  textColor: '--office-file-text-color',
  mutedTextColor: '--office-file-muted-text-color',
  disabledTextColor: '--office-file-disabled-text-color',
  surfaceColor: '--office-file-surface-color',
  subtleSurfaceColor: '--office-file-subtle-surface-color',
  workspaceColor: '--office-file-workspace-color',
  selectedSurfaceColor: '--office-file-selected-surface-color',
  borderColor: '--office-file-border-color',
  controlBorderColor: '--office-file-control-border-color',
  disabledSurfaceColor: '--office-file-disabled-surface-color',
  primaryHoverColor: '--office-file-primary-hover-color',
  primaryActiveColor: '--office-file-primary-active-color',
  focusRingColor: '--office-file-focus-ring-color',
  overlayShadow: '--office-file-overlay-shadow',
} as const;

/** 解析三位或六位十六进制颜色，供主色状态色自动推导。 */
function parseHexColor(color: string): [number, number, number] | undefined {
  const normalized = color.trim();
  const shortMatch = /^#([\da-f])([\da-f])([\da-f])$/i.exec(normalized);
  if (shortMatch) {
    return shortMatch
      .slice(1)
      .map((value) => Number.parseInt(value + value, 16)) as [
      number,
      number,
      number,
    ];
  }
  const longMatch = /^#([\da-f]{2})([\da-f]{2})([\da-f]{2})$/i.exec(normalized);
  if (!longMatch) return undefined;
  return longMatch.slice(1).map((value) => Number.parseInt(value, 16)) as [
    number,
    number,
    number,
  ];
}

/** 将主色向目标色混合，非十六进制颜色则保持原值以确保 CSS 始终有效。 */
function mixThemeColor(
  color: string,
  target: readonly [number, number, number],
  ratio: number,
): string {
  const source = parseHexColor(color);
  if (!source) return color;
  const channels = source.map((value, index) =>
    Math.round(value + (target[index] - value) * ratio),
  );
  return `#${channels
    .map((value) => value.toString(16).padStart(2, '0'))
    .join('')}`;
}

/** 将十六进制主色转换为带透明度的焦点环颜色。 */
function resolveFocusRingColor(color: string): string {
  const channels = parseHexColor(color);
  return channels ? `rgb(${channels.join(' ')} / 24%)` : color;
}

/** 取得主题模式，并为未配置主题保持原有浅色外观。 */
export function resolveOfficeViewerThemeMode(
  theme: OfficeViewerThemeOptions | undefined,
): OfficeViewerThemeMode {
  return theme?.mode ?? 'light';
}

/** 把公开主题配置转换成稳定的外壳 CSS 变量。 */
export function resolveOfficeViewerThemeStyle(
  theme: OfficeViewerThemeOptions | undefined,
): OfficeViewerThemeStyle {
  const style: OfficeViewerThemeStyle = {};
  const primaryColor = theme?.primaryColor;
  if (primaryColor) {
    style['--office-file-primary-color'] = primaryColor;
    style['--office-file-primary-hover-color'] = mixThemeColor(
      primaryColor,
      [255, 255, 255],
      0.2,
    );
    style['--office-file-primary-active-color'] = mixThemeColor(
      primaryColor,
      [0, 0, 0],
      0.18,
    );
    style['--office-file-focus-ring-color'] =
      resolveFocusRingColor(primaryColor);
  }

  const tokens = theme?.tokens;
  if (!tokens) return style;
  Object.entries(THEME_TOKEN_VARIABLES).forEach(([tokenName, variableName]) => {
    const value = tokens[tokenName as keyof typeof THEME_TOKEN_VARIABLES];
    if (value !== undefined) style[variableName] = value;
  });
  if (tokens.controlRadius !== undefined) {
    style['--office-file-radius'] =
      typeof tokens.controlRadius === 'number'
        ? `${tokens.controlRadius}px`
        : tokens.controlRadius;
  }
  return style;
}
