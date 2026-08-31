import React, { memo, type ReactNode } from 'react';
import { SpreadsheetGridPlaceholder } from './SpreadsheetGridPlaceholder';

/** 只覆盖 Sheet 内容区，保留 Tabs 和工具栏的状态边界。 */
function SpreadsheetSheetStateComponent({
  loading,
  sheetName,
  error,
  retry,
  children,
}: {
  loading: boolean;
  /** 当前工作表名称，用于生成上下文加载提示。 */
  sheetName?: string;
  error?: Error;
  retry(): void;
  children?: ReactNode;
}) {
  if (error) {
    return (
      <SpreadsheetGridPlaceholder
        error={error}
        sheetName={sheetName}
        onRetry={retry}
      />
    );
  }
  if (loading)
    return <SpreadsheetGridPlaceholder loading sheetName={sheetName} />;
  return <>{children}</>;
}

export const SpreadsheetSheetState = memo(SpreadsheetSheetStateComponent);
