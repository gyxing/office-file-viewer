import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import { OfficeButton } from '../../shared/ui/OfficeButton';
import { OfficeSpinner } from '../../shared/ui/OfficeSpinner';

/** 定义大工作表范围占位组件的属性。 */
type SpreadsheetGridPlaceholderProps = {
  /** 当前是否仍在加载。 */
  loading?: boolean;
  /** 当前工作表名称，用于生成上下文加载提示。 */
  sheetName?: string;
  /** 当前范围读取错误。 */
  error?: Error;
  /** 重试当前 Sheet 或范围。 */
  onRetry?: () => void;
};

/** 渲染工作表或虚拟范围的本地 loading/error 状态。 */
function SpreadsheetGridPlaceholderComponent({
  loading,
  sheetName,
  error,
  onRetry,
}: SpreadsheetGridPlaceholderProps) {
  const messages = useOfficeFileViewerMessages();
  const loadingLabel = sheetName
    ? messages.lazyContent.sheetLoading(sheetName)
    : messages.lazyContent.loading;
  return (
    <div className="office-file-xlsx-grid-placeholder">
      {loading ? (
        <>
          <OfficeSpinner label={loadingLabel} />
          <span>{loadingLabel}</span>
        </>
      ) : (
        <>
          <span>
            {messages.lazyContent.sheetLoadFailed}
            {error?.message ? `：${error.message}` : ''}
          </span>
          {onRetry ? (
            <OfficeButton size="small" onClick={onRetry}>
              {messages.lazyContent.retry}
            </OfficeButton>
          ) : null}
        </>
      )}
    </div>
  );
}

export const SpreadsheetGridPlaceholder = memo(
  SpreadsheetGridPlaceholderComponent,
);
