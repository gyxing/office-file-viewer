import type { ComponentType } from 'react';
import type { OfficeDocumentRuntime } from '../core/types';

/** Viewer Renderer 接收的公开运行时上下文。 */
export type OfficeViewerDocumentProps<TModel = unknown, TSource = unknown> = {
  /** 当前文件的解析运行时。 */
  runtime: OfficeDocumentRuntime<TModel, TSource>;
  /** 当前文件名。 */
  fileName: string;
  /** Renderer 内部错误回调。 */
  onError?: (error: unknown) => void;
};

/**
 * 将 Core 插件映射到 Viewer React 内容的适配器。
 * @experimental Viewer 插件适配协议首版。
 */
export type OfficeViewerPluginAdapter<TModel = unknown, TSource = unknown> = {
  /** 必须与 Core 插件 ID 完全一致。 */
  pluginId: string;
  /** 渲染插件文档内容。 */
  renderDocument: ComponentType<OfficeViewerDocumentProps<TModel, TSource>>;
  /** 可选的插件专属不支持状态。 */
  renderUnsupported?: ComponentType<{ fileName: string }>;
};
