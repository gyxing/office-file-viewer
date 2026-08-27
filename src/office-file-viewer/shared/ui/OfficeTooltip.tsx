import type { KeyboardEvent, ReactElement } from 'react';
import React, { cloneElement, useState } from 'react';
import './index.less';

/** Tooltip 支持的工具栏定位方式。 */
export type OfficeTooltipPlacement = 'bottom' | 'bottom-end';

/** 预览器内部提示属性。 */
export type OfficeTooltipProps = {
  /** 提示中展示的本地化文字。 */
  content: string;
  /** 气泡相对触发元素的定位方式。 */
  placement?: OfficeTooltipPlacement;
  /** 接收提示语义的单个触发元素。 */
  children: ReactElement;
};

/** 用于为 React 16 环境生成稳定且不依赖 useId 的提示标识。 */
let tooltipIdSequence = 0;

/** 生成当前运行时内唯一的提示标识。 */
function createTooltipId() {
  tooltipIdSequence += 1;
  return `office-file-tooltip-${tooltipIdSequence}`;
}

/** 渲染无需全局弹层容器的内部提示。 */
export function OfficeTooltip({
  children,
  content,
  placement = 'bottom',
}: OfficeTooltipProps) {
  const [dismissed, setDismissed] = useState(false);
  const [tooltipId] = useState(createTooltipId);
  const existingDescription = children.props['aria-describedby'];
  const describedBy = [existingDescription, tooltipId]
    .filter(Boolean)
    .join(' ');
  const describedChild = cloneElement(children, {
    'aria-describedby': describedBy,
  });

  const handleKeyDown = (event: KeyboardEvent<HTMLSpanElement>) => {
    if (event.key === 'Escape') setDismissed(true);
  };

  const handleClick = () => {
    // 原生按钮点击后仍会保留焦点，需要覆盖 focus-within，避免气泡持续显示。
    setDismissed(true);
  };

  return (
    <span
      className={[
        'office-file-tooltip',
        `office-file-tooltip--${placement}`,
      ].join(' ')}
      data-dismissed={dismissed ? 'true' : 'false'}
      onBlurCapture={() => setDismissed(false)}
      onClickCapture={handleClick}
      onFocusCapture={() => setDismissed(false)}
      onKeyDownCapture={handleKeyDown}
      onPointerEnter={() => setDismissed(false)}
      onPointerLeave={() => setDismissed(true)}
    >
      {describedChild}
      <span
        id={tooltipId}
        className="office-file-tooltip__bubble"
        role="tooltip"
      >
        {content}
      </span>
    </span>
  );
}
