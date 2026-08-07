import type { ButtonHTMLAttributes, ReactNode } from 'react';
import React from 'react';
import './index.less';

/** 内部按钮支持的视觉强调级别。 */
export type OfficeButtonVariant = 'default' | 'primary';

/** 内部按钮支持的紧凑尺寸。 */
export type OfficeButtonSize = 'default' | 'small';

/** 预览器内部按钮属性。 */
export type OfficeButtonProps = Omit<
  ButtonHTMLAttributes<HTMLButtonElement>,
  'type'
> & {
  /** 按钮的视觉强调级别。 */
  variant?: OfficeButtonVariant;
  /** 按钮的尺寸。 */
  size?: OfficeButtonSize;
  /** 显示在文字前方的图标。 */
  icon?: ReactNode;
};

/** 渲染预览器内部使用的原生按钮。 */
export function OfficeButton({
  children,
  className,
  icon,
  size = 'default',
  variant = 'default',
  ...buttonProps
}: OfficeButtonProps) {
  const hasLabel =
    children !== undefined && children !== null && children !== false;
  const mergedClassName = [
    'office-file-button',
    `office-file-button--${variant}`,
    size === 'small' ? 'office-file-button--small' : undefined,
    hasLabel ? undefined : 'office-file-button--icon-only',
    className,
  ]
    .filter(Boolean)
    .join(' ');

  return (
    <button {...buttonProps} type="button" className={mergedClassName}>
      {icon ? (
        <span className="office-file-button__icon" aria-hidden="true">
          {icon}
        </span>
      ) : null}
      {hasLabel ? (
        <span className="office-file-button__label">{children}</span>
      ) : null}
    </button>
  );
}
