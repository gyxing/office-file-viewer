import type { CSSProperties } from 'react';

/** 预览外壳支持的配色模式。 */
export type OfficeViewerThemeMode = 'light' | 'dark' | 'system';

/** 允许宿主覆盖的稳定语义色与外观令牌。 */
export type OfficeViewerThemeTokens = Readonly<{
  /** 主要文字颜色。 */
  textColor?: string;
  /** 次要说明文字颜色。 */
  mutedTextColor?: string;
  /** 禁用状态文字颜色。 */
  disabledTextColor?: string;
  /** 工具栏、侧栏和浮层的基础背景色。 */
  surfaceColor?: string;
  /** 次级区域与禁用控件的浅层背景色。 */
  subtleSurfaceColor?: string;
  /** 文档工作区背景色。 */
  workspaceColor?: string;
  /** 选中项背景色。 */
  selectedSurfaceColor?: string;
  /** 面板与区域分隔线颜色。 */
  borderColor?: string;
  /** 输入框、按钮等控件边框颜色。 */
  controlBorderColor?: string;
  /** 禁用控件背景色。 */
  disabledSurfaceColor?: string;
  /** 主色悬停状态颜色。 */
  primaryHoverColor?: string;
  /** 主色按下状态颜色。 */
  primaryActiveColor?: string;
  /** 键盘焦点环颜色。 */
  focusRingColor?: string;
  /** 浮层阴影。 */
  overlayShadow?: string;
  /** 控件圆角；数字按像素处理。 */
  controlRadius?: string | number;
}>;

/** 预览器或复用外壳的主题配置。 */
export type OfficeViewerThemeOptions = Readonly<{
  /** 配色模式，默认使用浅色。 */
  mode?: OfficeViewerThemeMode;
  /** 工具栏按钮、焦点和选中状态使用的主题主色。 */
  primaryColor?: string;
  /** 覆盖外壳语义令牌，不影响源文档正文样式。 */
  tokens?: OfficeViewerThemeTokens;
}>;

/** 主题转换后写入根节点的 CSS 变量。 */
export interface OfficeViewerThemeStyle extends CSSProperties {
  '--office-file-primary-color'?: string;
  '--office-file-primary-hover-color'?: string;
  '--office-file-primary-active-color'?: string;
  '--office-file-text-color'?: string;
  '--office-file-muted-text-color'?: string;
  '--office-file-disabled-text-color'?: string;
  '--office-file-surface-color'?: string;
  '--office-file-subtle-surface-color'?: string;
  '--office-file-workspace-color'?: string;
  '--office-file-selected-surface-color'?: string;
  '--office-file-border-color'?: string;
  '--office-file-control-border-color'?: string;
  '--office-file-disabled-surface-color'?: string;
  '--office-file-focus-ring-color'?: string;
  '--office-file-overlay-shadow'?: string;
  '--office-file-radius'?: string;
}
