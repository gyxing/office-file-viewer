/** 可复用预览外壳的独立入口，避免宿主仅使用布局时引入格式解析器。 */
export type {
  OfficeViewerThemeMode,
  OfficeViewerThemeOptions,
  OfficeViewerThemeTokens,
} from './office-file-viewer/shared/theme';
export type {
  OfficeViewerWatermark,
  OfficeViewerWatermarkOptions,
} from './office-file-viewer/shared/watermark';
export {
  OfficeViewerLayout,
  useOfficeViewerLayout,
} from './office-file-viewer/shell/layout';
export type {
  OfficeViewerLayoutActions,
  OfficeViewerLayoutContentScaling,
  OfficeViewerLayoutContextValue,
  OfficeViewerLayoutMeta,
  OfficeViewerLayoutProps,
  OfficeViewerLayoutState,
} from './office-file-viewer/shell/layout';
