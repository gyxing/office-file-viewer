import type { ReactNode, RefObject } from 'react';
import React, { memo, useEffect, useRef } from 'react';
import './index.less';
import { useOfficeNavigationResize } from './useOfficeNavigationResize';

/** 通用左侧导航面板属性。 */
type OfficeNavigationPanelProps = {
  /** 当前是否显示面板。 */
  visible: boolean;
  /** 切换文档时用于恢复默认宽度的稳定标识。 */
  sessionKey: string;
  /** 面板区域的无障碍名称。 */
  ariaLabel: string;
  /** 宽度分隔条的无障碍名称。 */
  resizeLabel: string;
  /** 面板内容。 */
  children: ReactNode;
  /** 追加到通用面板根节点的类名。 */
  className?: string;
  /** 供烟测和格式视图读取的附加 data 属性。 */
  dataAttributes?: Record<string, string | number | undefined>;
};

/** 通用左侧导航宽度分隔条属性。 */
type OfficeNavigationResizeHandleProps = {
  /** 待调整宽度的面板元素。 */
  panelRef: RefObject<HTMLElement>;
  /** 切换文档时用于恢复默认宽度的稳定标识。 */
  sessionKey: string;
  /** 分隔条的无障碍名称。 */
  label: string;
};

/** 渲染支持指针与键盘操作的通用导航宽度分隔条。 */
function OfficeNavigationResizeHandle({
  panelRef,
  sessionKey,
  label,
}: OfficeNavigationResizeHandleProps) {
  const {
    width,
    maxWidth,
    minWidth,
    handleRef,
    handleKeyDown,
    handlePointerDown,
    handlePointerMove,
    handlePointerEnd,
  } = useOfficeNavigationResize(panelRef, sessionKey);

  return (
    <div
      ref={handleRef}
      className="office-file-navigation-panel__resize-handle"
      role="separator"
      aria-label={label}
      aria-orientation="vertical"
      aria-valuemin={minWidth}
      aria-valuemax={Math.round(maxWidth)}
      aria-valuenow={Math.round(width)}
      title={label}
      tabIndex={0}
      onKeyDown={handleKeyDown}
      onPointerDown={handlePointerDown}
      onPointerMove={handlePointerMove}
      onPointerUp={handlePointerEnd}
      onPointerCancel={handlePointerEnd}
      onLostPointerCapture={handlePointerEnd}
    >
      <span aria-hidden="true" />
    </div>
  );
}

/** 提供可动画显隐、可调整宽度且隐藏后不可聚焦的左侧导航壳。 */
function OfficeNavigationPanelComponent({
  visible,
  sessionKey,
  ariaLabel,
  resizeLabel,
  children,
  className,
  dataAttributes,
}: OfficeNavigationPanelProps) {
  const panelRef = useRef<HTMLElement>(null);

  useEffect(() => {
    const panel = panelRef.current;
    if (!panel) return;
    // 收起动画期间保留 DOM，但隐藏面板不能继续接收键盘焦点。
    if (visible) panel.removeAttribute('inert');
    else panel.setAttribute('inert', '');
  }, [visible]);

  return (
    <aside
      ref={panelRef}
      className={['office-file-navigation-panel', className]
        .filter(Boolean)
        .join(' ')}
      aria-label={ariaLabel}
      aria-hidden={!visible}
      data-visible={visible ? 'true' : 'false'}
      {...dataAttributes}
    >
      {children}
      <OfficeNavigationResizeHandle
        panelRef={panelRef}
        sessionKey={sessionKey}
        label={resizeLabel}
      />
    </aside>
  );
}

export const OfficeNavigationPanel = memo(OfficeNavigationPanelComponent);
