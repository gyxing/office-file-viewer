// OfficeError 展示文件解析或预览过程中的错误信息。
import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../locale';
import { OfficeNotice } from './Notice';

/** 定义 OfficeError 组件可接收的属性。 */
type OfficeErrorProps = {
  /** OfficeErrorProps 面向调用方或用户展示的具体警告、错误说明。 */
  message: string;
};

/** 渲染 OfficeErrorComponent 组件。 */
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
