import React from 'react';
import './index.less';

/** 内部加载指示器支持的尺寸。 */
export type OfficeSpinnerSize = 'default' | 'large';

/** 预览器内部加载指示器属性。 */
export type OfficeSpinnerProps = {
  /** 加载指示器的尺寸。 */
  size?: OfficeSpinnerSize;
  /** 辅助技术读取的加载状态文案。 */
  label?: string;
};

/** 渲染不依赖组件库的加载指示器。 */
export function OfficeSpinner({ label, size = 'default' }: OfficeSpinnerProps) {
  return (
    <span
      className={[
        'office-file-spinner',
        size === 'large' ? 'office-file-spinner--large' : undefined,
      ]
        .filter(Boolean)
        .join(' ')}
      aria-label={label}
      role="status"
    >
      <span className="office-file-spinner__indicator" aria-hidden="true" />
    </span>
  );
}
