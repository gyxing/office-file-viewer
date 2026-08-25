import type { ReactElement, ReactNode } from 'react';
import React, { createContext, useContext } from 'react';
import type { OfficeViewerLayoutContextValue } from './types';

const OfficeViewerLayoutContext = createContext<
  OfficeViewerLayoutContextValue | undefined
>(undefined);

/** 外壳上下文注入属性。 */
type OfficeViewerLayoutProviderProps = {
  /** 当前实例的状态、动作和环境信息。 */
  value: OfficeViewerLayoutContextValue;
  /** 需要消费外壳能力的内容。 */
  children: ReactNode;
};

/** 将控制器实现与工具栏、宿主内容解耦。 */
export function OfficeViewerLayoutProvider({
  value,
  children,
}: OfficeViewerLayoutProviderProps): ReactElement {
  return (
    <OfficeViewerLayoutContext.Provider value={value}>
      {children}
    </OfficeViewerLayoutContext.Provider>
  );
}

/** 读取最近一层可复用预览外壳的状态和操作。 */
export function useOfficeViewerLayout(): OfficeViewerLayoutContextValue {
  const context = useContext(OfficeViewerLayoutContext);
  if (!context) {
    throw new Error(
      'useOfficeViewerLayout 必须在 OfficeViewerLayout 内部使用。',
    );
  }
  return context;
}
