import type { ReactElement } from 'react';
import React, { useEffect, useMemo, useRef } from 'react';
import { createOfficePluginRegistry } from './OfficePluginRegistry';
import {
  OfficeViewerPluginContextProvider,
  type OfficeViewerPluginContextValue,
} from './OfficeViewerPluginContext';
import type { OfficeCorePlugin, OfficePluginRegistry } from './types';
import type { OfficeViewerPluginAdapter } from './viewerTypes';

/**
 * Viewer 插件 Provider 的公开属性。
 * @experimental 插件 Provider 首版。
 */
export type OfficeViewerProviderProps = {
  /** 宿主自行管理的注册表；传入后不会由 Provider 释放。 */
  registry?: OfficePluginRegistry;
  /** 需要在 Provider 作用域中注册的插件。 */
  plugins?: readonly OfficeCorePlugin[];
  /** 插件 ID 到 React Renderer 的适配器。 */
  adapters?: readonly OfficeViewerPluginAdapter[];
  /** 需要接收插件上下文的内容。 */
  children: React.ReactNode;
};

/** 为 Viewer 建立作用域级插件注册和渲染边界。 */
export function OfficeViewerProvider({
  registry,
  plugins = [],
  adapters = [],
  children,
}: OfficeViewerProviderProps): ReactElement {
  const ownedRegistryRef = useRef<OfficePluginRegistry>();
  const lifecycleTokenRef = useRef<symbol>();
  if (!registry && !ownedRegistryRef.current) {
    ownedRegistryRef.current = createOfficePluginRegistry({
      includeBuiltIns: true,
      plugins,
    });
  }
  const activeRegistry = registry ?? ownedRegistryRef.current!;
  const previousPluginsRef = useRef<readonly OfficeCorePlugin[]>(plugins);

  useEffect(() => {
    if (
      registry &&
      plugins.length > 0 &&
      typeof process !== 'undefined' &&
      process.env.NODE_ENV !== 'production'
    ) {
      console.warn(
        'OfficeViewerProvider 同时收到 registry 和 plugins；plugins 将被忽略。',
      );
    }
  }, [plugins.length, registry]);

  useEffect(() => {
    if (registry) {
      // 外部 Registry 接管期间 plugins 不参与注册，但仍记录最新快照，
      // 这样切回非受控模式时不会把已在新 Registry 初始注册的插件重复注册。
      previousPluginsRef.current = plugins;
      return undefined;
    }
    const previous = new Map(
      previousPluginsRef.current.map((plugin) => [plugin.id, plugin]),
    );
    const current = new Map(plugins.map((plugin) => [plugin.id, plugin]));
    current.forEach((plugin, id) => {
      if (previous.get(id) === plugin) return;
      if (previous.has(id)) activeRegistry.unregister(id);
      activeRegistry.register(plugin);
    });
    previous.forEach((plugin, id) => {
      if (current.has(id)) return;
      activeRegistry.unregister(id);
      void plugin;
    });
    previousPluginsRef.current = plugins;
    return undefined;
  }, [activeRegistry, plugins, registry]);

  useEffect(() => {
    // Registry 属性可能在组件存续期间切换；延迟释放只在真正卸载时执行，
    // 避免切回非受控模式后复用一个已经释放的实例。微任务也兼容 React StrictMode
    // 的 effect 重放，不会在一次重放间隙误释放仍将继续使用的注册表。
    const lifecycleToken = Symbol('office-viewer-provider-lifecycle');
    lifecycleTokenRef.current = lifecycleToken;
    return () => {
      void Promise.resolve().then(() => {
        if (lifecycleTokenRef.current !== lifecycleToken) return;
        ownedRegistryRef.current?.dispose();
      });
    };
  }, []);

  const adapterMap = useMemo(() => {
    const map = new Map<string, OfficeViewerPluginAdapter>();
    adapters.forEach((adapter) => map.set(adapter.pluginId, adapter));
    return map;
  }, [adapters]);
  const value = useMemo<OfficeViewerPluginContextValue>(
    () => ({
      registry: activeRegistry,
      getAdapter: (pluginId) => adapterMap.get(pluginId),
    }),
    [activeRegistry, adapterMap],
  );

  return (
    <OfficeViewerPluginContextProvider value={value}>
      {children}
    </OfficeViewerPluginContextProvider>
  );
}
