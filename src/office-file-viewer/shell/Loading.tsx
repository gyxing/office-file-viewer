import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../locale';
import { OfficeSpinner } from '../shared/ui/OfficeSpinner';

/** Office 文件解析加载组件属性。 */
type OfficeLoadingProps = {
  /** 加载指示器下方展示的提示文字。 */
  tip?: string;
};

/** 展示 Office 文件解析中的统一加载状态。 */
function OfficeLoadingComponent({ tip }: OfficeLoadingProps) {
  const messages = useOfficeFileViewerMessages();
  const label = tip ?? messages.loading.parsing;

  return (
    <div className="office-file-loading">
      <OfficeSpinner size="large" label={label} />
      <span className="office-file-loading__tip">{label}</span>
    </div>
  );
}

export const OfficeLoading = memo(OfficeLoadingComponent);
