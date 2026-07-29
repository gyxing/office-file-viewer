// OfficeLoading 展示文件解析中的统一加载状态。
import { Spin } from 'antd';
import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../locale';

/** 定义 OfficeLoading 组件可接收的属性。 */
type OfficeLoadingProps = {
  /** OfficeLoadingProps 的 tip 文本值。 */
  tip?: string;
};

/** 渲染 OfficeLoadingComponent 组件。 */
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
