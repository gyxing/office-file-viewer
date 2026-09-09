import type { ReactElement } from 'react';
import React, { createContext, useContext } from 'react';
import type { OfficePluginResolver } from '../core/parsingContracts';
import { createOfficePluginRegistry } from './OfficePluginRegistry';
import type { OfficePluginRegistry } from './types';
import type { OfficeViewerPluginAdapter } from './viewerTypes';

/** Viewer 内部使用的插件注册和渲染适配器上下文。 */
export type OfficeViewerPluginContextValue = Readonly<{
  /** 当前作用域的 Core 插件注册表。 */
  registry: OfficePluginRegistry;
  /** 按插件 ID 查询 Viewer Renderer。 */
  getAdapter(pluginId: string): OfficeViewerPluginAdapter | undefined;
}>;

const OfficeViewerPluginContext = createContext<
  OfficeViewerPluginContextValue | undefined
>(undefined);

let defaultOfficeViewerPluginResolver: OfficePluginResolver | undefined;

/** 惰性创建只包含内置格式的默认注册表，保证普通 Viewer 也沿用同一识别边界。 */
export function getDefaultOfficeViewerPluginResolver() {
  if (!defaultOfficeViewerPluginResolver) {
    const registry = createOfficePluginRegistry({ includeBuiltIns: true });
    defaultOfficeViewerPluginResolver = { resolve: registry.resolve };
  }
  return defaultOfficeViewerPluginResolver;
}

/** 读取最近一层 Viewer 插件上下文；未包裹 Provider 时返回 undefined。 */
export function useOfficeViewerPluginContext() {
  return useContext(OfficeViewerPluginContext);
}

/** 向 Viewer 子树提供实例级插件注册表和 Renderer Adapter。 */
export function OfficeViewerPluginContextProvider({
  value,
  children,
}: {
  /** 当前作用域的插件上下文。 */
  value: OfficeViewerPluginContextValue;
  /** 需要消费插件的 Viewer 子树。 */
  children: React.ReactNode;
}): ReactElement {
  return (
    <OfficeViewerPluginContext.Provider value={value}>
      {children}
    </OfficeViewerPluginContext.Provider>
  );
}
