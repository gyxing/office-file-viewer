// OfficeNotice 展示预览器统一的错误或警告状态。
import React, { memo } from 'react';

/** Office提示组件属性。 */
type OfficeNoticeProps = {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'error' | 'warning';
  /** 面向用户展示的标题。 */
  title: string;
  /** 补充说明内容。 */
  description: string;
};

/** 展示不阻断预览流程的提示信息。 */
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
