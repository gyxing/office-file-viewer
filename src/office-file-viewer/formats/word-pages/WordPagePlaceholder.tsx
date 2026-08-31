import React from 'react';
import { useOfficeFileViewerMessages } from '../../locale';

/** Word 页面加载占位组件属性。 */
type WordPagePlaceholderProps = {
  /** 当前加载或解析状态。 */
  status: 'loading' | 'error';
  /** 当前占位页面的显示序号。 */
  pageNumber?: number;
  /** 最小高度，单位为标准化渲染像素。 */
  minHeight: number;
  /** 在用户请求重试当前内容时触发。 */
  onRetry?(): void;
};

/** 显示按需页面的本地化加载或失败状态。 */
export function WordPagePlaceholder({
  status,
  pageNumber,
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
          : pageNumber === undefined
          ? messages.lazyContent.loading
          : messages.lazyContent.pageLoading(pageNumber)}
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
