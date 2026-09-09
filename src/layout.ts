/** 可复用预览外壳的独立入口，避免宿主仅使用布局时引入格式解析器。 */
export type {
  OfficeViewerThemeMode,
  OfficeViewerThemeOptions,
  OfficeViewerThemeTokens,
} from './office-file-viewer/shared/theme/index';
export type {
  OfficeViewerWatermark,
  OfficeViewerWatermarkOptions,
} from './office-file-viewer/shared/watermark/index';
export {
  OfficeViewerLayout,
  OfficeViewerShell,
  useOfficeViewerLayout,
} from './office-file-viewer/shell/layout/index';
export type {
  OfficeViewerLayoutActions,
  OfficeViewerLayoutContentScaling,
  OfficeViewerLayoutContextValue,
  OfficeViewerLayoutMeta,
  OfficeViewerLayoutProps,
  OfficeViewerLayoutState,
  OfficeViewerShellComponent,
  OfficeViewerShellContextValue,
  OfficeViewerShellProviderProps,
} from './office-file-viewer/shell/layout/index';
