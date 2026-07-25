// OfficeEmpty 按文件格式展示空状态，提示用户上传对应 Office 文件。
import { Empty } from 'antd';
import React, { memo } from 'react';
import type { PreviewKind } from '../services/preview';
import { OFFICE_EMPTY_DESCRIPTIONS } from './constants';

/** 定义 OfficeEmpty 组件可接收的属性。 */
type OfficeEmptyProps = {
  /** 标识 OfficeEmptyProps 对应的 Office 文件或数据种类。 */
  kind: PreviewKind;
};

/** 渲染 OfficeEmptyComponent 组件。 */
function OfficeEmptyComponent({ kind }: OfficeEmptyProps) {
  return <Empty description={OFFICE_EMPTY_DESCRIPTIONS[kind]} />;
}

export const OfficeEmpty = memo(OfficeEmptyComponent);
