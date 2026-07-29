import type { ReactNode } from 'react';
import React, { createContext, useContext } from 'react';
import type { OfficeResourceStore } from './types';

const OfficeResourceContext = createContext<OfficeResourceStore | undefined>(
  undefined,
);

type OfficeResourceStoreProviderProps = {
  store: OfficeResourceStore;
  children?: ReactNode;
};

/** 为单个文档预览树提供隔离的资源 Store。 */
export function OfficeResourceStoreProvider({
  store,
  children,
}: OfficeResourceStoreProviderProps) {
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
