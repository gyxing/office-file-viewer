import { Empty } from 'antd';
import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../locale';
import type { PreviewKind } from '../services/preview';

/** Office 空状态组件属性。 */
type OfficeEmptyProps = {
  /** 已识别的内容格式；未提供表示尚未选择文件。 */
  kind?: PreviewKind;
};

/** 展示未选择文件或格式内容为空时的提示。 */
function OfficeEmptyComponent({ kind }: OfficeEmptyProps) {
  const messages = useOfficeFileViewerMessages();
  const description = kind
    ? messages.empty[kind]
    : messages.file.selectToPreview;
  return <Empty description={description} />;
}

export const OfficeEmpty = memo(OfficeEmptyComponent);
