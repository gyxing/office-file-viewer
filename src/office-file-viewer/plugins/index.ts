/** 插件公共入口，按需导出注册表、类型和 Viewer Provider。 */
export { createOfficePluginRegistry } from './OfficePluginRegistry';
export { OfficeViewerProvider } from './OfficeViewerProvider';
export type { OfficeViewerProviderProps } from './OfficeViewerProvider';
export type {
  OfficeCorePlugin,
  OfficeFormatPlugin,
  OfficePluginRegistry,
  OfficePluginRegistryOptions,
  OfficePluginWorkerSupport,
} from './types';
export type {
  OfficeViewerDocumentProps,
  OfficeViewerPluginAdapter,
} from './viewerTypes';
