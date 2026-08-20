import type { ReactNode } from 'react';
import React, { createContext, useContext, useEffect } from 'react';
import type { OfficeAnnotationSource } from '../../services/annotations/AnnotationSource';
import type { OfficeAnnotationTarget } from '../../services/annotations/types';
import type {
  OfficeAnnotationController,
  OfficeAnnotationNavigator,
} from './useOfficeAnnotationController';

const OfficeAnnotationRuntimeContext = createContext<
  OfficeAnnotationController | undefined
>(undefined);

/** 为审阅面板和各格式渲染器提供同一个批注控制器。 */
export function OfficeAnnotationProvider({
  controller,
  children,
}: {
  /** 当前查看器独占的批注控制器。 */
  controller: OfficeAnnotationController;
  /** 工具栏、审阅面板与格式查看器。 */
  children: ReactNode;
}) {
  return (
    <OfficeAnnotationRuntimeContext.Provider value={controller}>
      {children}
    </OfficeAnnotationRuntimeContext.Provider>
  );
}

/** 仅在审阅能力启用时建立运行时，避免关闭能力后注册额外数据源。 */
export function OfficeAnnotationRuntimeBoundary({
  enabled,
  controller,
  children,
}: {
  /** 当前查看器是否启用审阅能力。 */
  enabled: boolean;
  /** 当前查看器独占的批注控制器。 */
  controller: OfficeAnnotationController;
  /** 当前查看器的完整渲染内容。 */
  children: ReactNode;
}) {
  if (!enabled) return <>{children}</>;
  return (
    <OfficeAnnotationProvider controller={controller}>
      {children}
    </OfficeAnnotationProvider>
  );
}

/** 读取当前查看器的审阅运行时；未启用时返回空值。 */
export function useOfficeAnnotationRuntime() {
  return useContext(OfficeAnnotationRuntimeContext);
}

/** 在格式查看器挂载期间注册其批注数据源。 */
export function useOfficeAnnotationSourceRegistration(
  source: OfficeAnnotationSource | undefined,
) {
  const runtime = useOfficeAnnotationRuntime();
  const registerSource = runtime?.actions.registerSource;
  useEffect(() => {
    if (!source || !registerSource) return undefined;
    return registerSource(source);
  }, [registerSource, source]);
}

/** 在格式查看器挂载期间注册批注精确定位能力。 */
export function useOfficeAnnotationNavigatorRegistration(
  kind: OfficeAnnotationTarget['kind'],
  navigator: OfficeAnnotationNavigator | undefined,
) {
  const runtime = useOfficeAnnotationRuntime();
  const registerNavigator = runtime?.actions.registerNavigator;
  useEffect(() => {
    if (!navigator || !registerNavigator) return undefined;
    return registerNavigator(kind, navigator);
  }, [kind, navigator, registerNavigator]);
}
