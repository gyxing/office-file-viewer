// OfficeLoading 展示文件解析中的统一加载状态。
import { Spin } from 'antd';
import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../locale';

/** Office 文件解析加载组件属性。 */
type OfficeLoadingProps = {
  /** 加载指示器下方展示的提示文字。 */
  tip?: string;
};

/** 展示 Office 文件解析中的统一加载状态。 */
function OfficeLoadingComponent({ tip }: OfficeLoadingProps) {
  const messages = useOfficeFileViewerMessages();
  return (
    <div className="office-file-loading">
      <Spin size="large" tip={tip ?? messages.loading.parsing}>
        <div className="office-file-loading__placeholder" />
      </Spin>
    </div>
  );
}

export const OfficeLoading = memo(OfficeLoadingComponent);
