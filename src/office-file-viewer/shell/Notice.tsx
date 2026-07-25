// OfficeNotice 展示跨 antd 版本稳定的错误或警告状态。
import React, { memo } from 'react';

/** 定义 OfficeNotice 组件可接收的属性。 */
type OfficeNoticeProps = {
  /** 用于区分 OfficeNoticeProps 不同结构分支的类型标识。 */
  type: 'error' | 'warning';
  /** OfficeNoticeProps 对外展示的标题。 */
  title: string;
  /** OfficeNoticeProps 的 description 文本值。 */
  description: string;
};

/** 渲染 OfficeNoticeComponent 组件。 */
function OfficeNoticeComponent({
  type,
  title,
  description,
}: OfficeNoticeProps) {
  return (
    <div
      className={`office-file-notice office-file-notice--${type}`}
      role="alert"
    >
      <span className="office-file-notice__icon" aria-hidden="true">
        !
      </span>
      <div className="office-file-notice__content">
        <div className="office-file-notice__title">{title}</div>
        <div className="office-file-notice__description">{description}</div>
      </div>
    </div>
  );
}

export const OfficeNotice = memo(OfficeNoticeComponent);
