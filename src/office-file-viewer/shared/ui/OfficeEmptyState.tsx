import React from 'react';
import { OfficeEmptyIcon } from './OfficeEmptyIcon';
import './index.less';

/** 预览器内部空状态属性。 */
export type OfficeEmptyStateProps = {
  /** 空状态展示的本地化说明。 */
  description: string;
  /** 调用位置附加的样式类名。 */
  className?: string;
};

/** 展示不指定 Office 类型的通用空状态。 */
export function OfficeEmptyState({
  className,
  description,
}: OfficeEmptyStateProps) {
  return (
    <div
      className={['office-file-empty-state', className]
        .filter(Boolean)
        .join(' ')}
      role="status"
    >
      <OfficeEmptyIcon className="office-file-empty-state__icon" />
      <span className="office-file-empty-state__description">
        {description}
      </span>
    </div>
  );
}
