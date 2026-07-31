import { Button, Spin } from 'antd';
import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';

/** 定义大工作表范围占位组件的属性。 */
type SpreadsheetGridPlaceholderProps = {
  /** 当前是否仍在加载。 */
  loading?: boolean;
  /** 当前范围读取错误。 */
  error?: Error;
  /** 重试当前 Sheet 或范围。 */
  onRetry?: () => void;
};

/** 渲染工作表或虚拟范围的本地 loading/error 状态。 */
function SpreadsheetGridPlaceholderComponent({
  loading,
  error,
  onRetry,
}: SpreadsheetGridPlaceholderProps) {
  const messages = useOfficeFileViewerMessages();
  return (
    <div className="office-file-xlsx-grid-placeholder">
      {loading ? (
        <>
          <Spin aria-label={messages.lazyContent.loading} />
          <span>{messages.lazyContent.loading}</span>
        </>
      ) : (
        <>
          <span>
            {messages.lazyContent.sheetLoadFailed}
            {error?.message ? `：${error.message}` : ''}
          </span>
          {onRetry ? (
            <Button size="small" onClick={onRetry}>
              {messages.lazyContent.retry}
            </Button>
          ) : null}
        </>
      )}
    </div>
  );
}

export const SpreadsheetGridPlaceholder = memo(
  SpreadsheetGridPlaceholderComponent,
);
