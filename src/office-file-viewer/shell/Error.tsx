// OfficeError 展示文件解析或预览过程中的错误信息。
import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../locale';
import { OfficeNotice } from './Notice';

/** Office错误组件属性。 */
type OfficeErrorProps = {
  /** 面向调用方或用户展示的说明。 */
  message: string;
};

/** 展示阻止继续预览的错误信息。 */
function OfficeErrorComponent({ message }: OfficeErrorProps) {
  const messages = useOfficeFileViewerMessages();
  return (
    <div className="office-file-error">
      <OfficeNotice
        type="error"
        title={messages.error.previewFailed}
        description={message}
      />
    </div>
  );
}

export const OfficeError = memo(OfficeErrorComponent);
