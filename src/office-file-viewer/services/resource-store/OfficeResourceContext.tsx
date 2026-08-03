import type { ReactElement, ReactNode } from 'react';
import React, { createContext, useContext } from 'react';
import type { OfficeResourceStore } from './types';

const OfficeResourceContext = createContext<OfficeResourceStore | undefined>(
  undefined,
);

/** Office 资源存储上下文组件属性。 */
type OfficeResourceStoreProviderProps = {
  /** 当前文档会话独占的资源存储。 */
  store: OfficeResourceStore;
  /** 当前组件包含的子节点。 */
  children?: ReactNode;
};

/** 为单个文档预览树提供隔离的资源 Store。 */
export function OfficeResourceStoreProvider({
  store,
  children,
}: OfficeResourceStoreProviderProps): ReactElement {
  return (
    <OfficeResourceContext.Provider value={store}>
      {children}
    </OfficeResourceContext.Provider>
  );
}

/** 读取当前文档会话的资源 Store。 */
export function useOfficeResourceStore() {
  return useContext(OfficeResourceContext);
}
