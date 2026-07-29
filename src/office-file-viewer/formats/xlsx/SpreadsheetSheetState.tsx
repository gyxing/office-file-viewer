import React, { memo, type ReactNode } from 'react';
import { SpreadsheetGridPlaceholder } from './SpreadsheetGridPlaceholder';

/** 只覆盖 Sheet 内容区，保留 Tabs 和工具栏的状态边界。 */
function SpreadsheetSheetStateComponent({
  loading,
  error,
  retry,
  children,
}: {
  loading: boolean;
  error?: Error;
  retry(): void;
  children?: ReactNode;
}) {
  if (error) {
    return <SpreadsheetGridPlaceholder error={error} onRetry={retry} />;
  }
  if (loading) return <SpreadsheetGridPlaceholder loading />;
  return <>{children}</>;
}

export const SpreadsheetSheetState = memo(SpreadsheetSheetStateComponent);
