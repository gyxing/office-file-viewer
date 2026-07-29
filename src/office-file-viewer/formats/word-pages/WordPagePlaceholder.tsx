import React from 'react';
import { useOfficeFileViewerMessages } from '../../locale';

type WordPagePlaceholderProps = {
  status: 'loading' | 'error';
  minHeight: number;
  onRetry?(): void;
};

/** 显示按需页面的本地化加载或失败状态。 */
export function WordPagePlaceholder({
  status,
  minHeight,
  onRetry,
}: WordPagePlaceholderProps) {
  const messages = useOfficeFileViewerMessages();
  return (
    <div
      className="office-file-word-pages__placeholder"
      style={{ minHeight }}
      role={status === 'error' ? 'alert' : 'status'}
      aria-live="polite"
    >
      <span>
        {status === 'error'
          ? messages.lazyContent.pageLoadFailed
          : messages.lazyContent.loading}
      </span>
      {status === 'error' && onRetry ? (
        <button
          type="button"
          className="office-file-word-pages__retry"
          onClick={onRetry}
        >
          {messages.lazyContent.retry}
        </button>
      ) : null}
    </div>
  );
}
