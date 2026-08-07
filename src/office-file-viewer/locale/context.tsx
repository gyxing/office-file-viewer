import type { ReactElement, ReactNode } from 'react';
import React, { createContext, useContext } from 'react';
import { enUS } from './en-US';
import type { OfficeFileViewerLocale, OfficeFileViewerMessages } from './types';
import { zhCN } from './zh-CN';

/** 文件预览器语言上下文组件属性。 */
type OfficeFileViewerLocaleProviderProps = {
  /** 当前预览器使用的内置界面语言。 */
  locale: OfficeFileViewerLocale;
  /** 共享当前语言资源的预览器内容。 */
  children?: ReactNode;
};

/** 内置语言映射只在组件内部使用，避免宿主语言环境影响预览器文案。 */
const OFFICE_FILE_VIEWER_LOCALES: Record<
  OfficeFileViewerLocale,
  OfficeFileViewerMessages
> = {
  'zh-CN': zhCN,
  'en-US': enUS,
};

const OfficeFileViewerLocaleContext = createContext(zhCN);

/** 为单个 OfficeFileViewer 实例提供对应语言资源。 */
export function OfficeFileViewerLocaleProvider({
  locale,
  children,
}: OfficeFileViewerLocaleProviderProps): ReactElement {
  return (
    <OfficeFileViewerLocaleContext.Provider
      value={OFFICE_FILE_VIEWER_LOCALES[locale]}
    >
      {children}
    </OfficeFileViewerLocaleContext.Provider>
  );
}

/** 读取当前 OfficeFileViewer 实例的界面语言资源。 */
export function useOfficeFileViewerMessages() {
  return useContext(OfficeFileViewerLocaleContext);
}
