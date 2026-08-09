import type { CSSProperties, KeyboardEvent } from 'react';
import React, { useEffect, useRef, useState } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import { CopyIcon, OpenLinkIcon } from '../ui/OfficeIcons';

/** 超链接右键菜单的显示参数。 */
type OfficeHyperlinkContextMenuProps = {
  /** 菜单相对预览器左侧的位置。 */
  x: number;
  /** 菜单相对预览器顶部的位置。 */
  y: number;
  /** 当前菜单是否用于文档内部跳转。 */
  internal: boolean;
  /** 外部链接最终可复制的安全地址。 */
  copyTarget?: string;
  /** 当前预览器所在的文档对象。 */
  ownerDocument: Document;
  /** 执行当前链接的默认打开或内部跳转。 */
  onOpen(): void;
  /** 关闭右键菜单。 */
  onClose(): void;
};

/** 优先使用 Clipboard API，并兼容不支持异步剪贴板的浏览器环境。 */
async function copyText(value: string, ownerDocument: Document) {
  const clipboard = ownerDocument.defaultView?.navigator.clipboard;
  if (clipboard?.writeText) {
    try {
      await clipboard.writeText(value);
      return;
    } catch {
      // 非安全上下文可能拒绝 Clipboard API，继续使用受控的文本域降级。
    }
  }

  const textarea = ownerDocument.createElement('textarea');
  textarea.value = value;
  textarea.setAttribute('readonly', '');
  textarea.style.position = 'fixed';
  textarea.style.top = '-9999px';
  textarea.style.opacity = '0';
  ownerDocument.body.appendChild(textarea);
  textarea.select();
  const copied = ownerDocument.execCommand('copy');
  textarea.remove();
  if (!copied) throw new Error('Clipboard write was rejected');
}

/** 渲染安全外部链接和文档内部链接共用的右键操作菜单。 */
export function OfficeHyperlinkContextMenu({
  x,
  y,
  internal,
  copyTarget,
  ownerDocument,
  onOpen,
  onClose,
}: OfficeHyperlinkContextMenuProps) {
  const messages = useOfficeFileViewerMessages();
  const menuRef = useRef<HTMLDivElement>(null);
  const [copying, setCopying] = useState(false);
  const [copyError, setCopyError] = useState(false);

  useEffect(() => {
    const menu = menuRef.current;
    const firstItem =
      menu?.querySelector<HTMLButtonElement>('[role="menuitem"]');
    try {
      firstItem?.focus({ preventScroll: true });
    } catch {
      firstItem?.focus();
    }
    const view = ownerDocument.defaultView;
    if (!view) return undefined;
    view.addEventListener('resize', onClose);
    view.addEventListener('scroll', onClose, true);
    return () => {
      view.removeEventListener('resize', onClose);
      view.removeEventListener('scroll', onClose, true);
    };
  }, [onClose, ownerDocument]);

  const handleKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      onClose();
      return;
    }
    if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const items = Array.from(
      menuRef.current?.querySelectorAll<HTMLButtonElement>(
        '[role="menuitem"]:not(:disabled)',
      ) ?? [],
    );
    if (!items.length) return;
    const activeIndex = items.indexOf(
      ownerDocument.activeElement as HTMLButtonElement,
    );
    const nextIndex =
      event.key === 'Home'
        ? 0
        : event.key === 'End'
        ? items.length - 1
        : event.key === 'ArrowDown'
        ? (activeIndex + 1 + items.length) % items.length
        : (activeIndex - 1 + items.length) % items.length;
    items[nextIndex].focus();
  };

  const handleCopy = async () => {
    if (!copyTarget || copying) return;
    setCopying(true);
    setCopyError(false);
    try {
      await copyText(copyTarget, ownerDocument);
      onClose();
    } catch {
      setCopyError(true);
    } finally {
      setCopying(false);
    }
  };

  return (
    <div
      className="office-file-hyperlink-context-menu-layer"
      onContextMenu={(event) => event.preventDefault()}
      onPointerDown={(event) => {
        if (event.target === event.currentTarget) onClose();
      }}
    >
      <div
        ref={menuRef}
        className="office-file-hyperlink-context-menu"
        role="menu"
        aria-label={messages.hyperlink.contextMenu}
        style={{ left: x, top: y } as CSSProperties}
        onKeyDown={handleKeyDown}
        onPointerDown={(event) => event.stopPropagation()}
      >
        <button
          type="button"
          className="office-file-hyperlink-context-menu__item"
          role="menuitem"
          onClick={onOpen}
        >
          <OpenLinkIcon />
          <span>
            {internal ? messages.hyperlink.jump : messages.hyperlink.open}
          </span>
        </button>
        {copyTarget ? (
          <button
            type="button"
            className="office-file-hyperlink-context-menu__item"
            role="menuitem"
            disabled={copying}
            onClick={handleCopy}
          >
            <CopyIcon />
            <span>{messages.hyperlink.copy}</span>
          </button>
        ) : null}
        {copyError ? (
          <div
            className="office-file-hyperlink-context-menu__error"
            role="alert"
          >
            {messages.hyperlink.copyFailed}
          </div>
        ) : null}
      </div>
    </div>
  );
}
