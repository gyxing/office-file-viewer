// OfficeError 展示文件解析或预览过程中的错误信息。
import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../locale';
import { OfficeButton } from '../shared/ui/OfficeButton';
import { OfficeNotice } from './Notice';

/** Office错误组件属性。 */
type OfficeErrorProps = {
  /** 面向调用方或用户展示的说明。 */
  message: string;
  /** 最近文件来源仍可用时触发重新加载。 */
  onRetry?: () => void;
};

/** 展示阻止继续预览的错误信息。 */
function OfficeErrorComponent({ message, onRetry }: OfficeErrorProps) {
  const messages = useOfficeFileViewerMessages();
  return (
    <div className="office-file-error">
      <OfficeNotice
        type="error"
        title={messages.error.previewFailed}
        description={message}
        action={
          onRetry ? (
            <OfficeButton variant="primary" onClick={onRetry}>
              {messages.lazyContent.retry}
            </OfficeButton>
          ) : undefined
        }
      />
    </div>
  );
}

export const OfficeError = memo(OfficeErrorComponent);
