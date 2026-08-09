import type { CSSProperties, KeyboardEvent } from 'react';
import React, { useEffect, useRef, useState } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import { DownloadIcon, PreviewIcon } from '../ui/OfficeIcons';
import { downloadOfficeImage } from './downloadOfficeImage';
import type { OfficeImagePreviewTarget } from './types';
import { useOfficeImagePreviewResource } from './useOfficeImagePreviewResource';

/** 图片右键菜单在预览器内的显示参数。 */
type OfficeImageContextMenuProps = {
  /** 当前右键操作关联的图片。 */
  target: OfficeImagePreviewTarget;
  /** 菜单相对预览器左侧的位置。 */
  x: number;
  /** 菜单相对预览器顶部的位置。 */
  y: number;
  /** 是否显示下载操作。 */
  download: boolean;
  /** 当前预览器所在的文档对象。 */
  ownerDocument: Document;
  /** 打开当前图片的预览层。 */
  onPreview(): void;
  /** 关闭右键菜单。 */
  onClose(): void;
};

/** 渲染只包含预览和下载操作的图片右键菜单。 */
export function OfficeImageContextMenu({
  target,
  x,
  y,
  download,
  ownerDocument,
  onPreview,
  onClose,
}: OfficeImageContextMenuProps) {
  const messages = useOfficeFileViewerMessages();
  const resource = useOfficeImagePreviewResource(download ? target : undefined);
  const menuRef = useRef<HTMLDivElement>(null);
  const [downloading, setDownloading] = useState(false);
  const [downloadError, setDownloadError] = useState<string>();

  useEffect(() => {
    const menu = menuRef.current;
    menu?.querySelector<HTMLButtonElement>('[role="menuitem"]')?.focus();
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

  const handleDownload = async () => {
    if (!resource.url || downloading) return;
    setDownloading(true);
    setDownloadError(undefined);
    try {
      await downloadOfficeImage(resource.url, target, ownerDocument);
      onClose();
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      setDownloadError(messages.imagePreview.downloadFailed(reason));
    } finally {
      setDownloading(false);
    }
  };
  const errorMessage =
    downloadError ??
    (resource.error ? messages.imagePreview.loadFailed : undefined);

  return (
    <div
      className="office-file-image-context-menu-layer"
      onPointerDown={(event) => {
        if (event.target === event.currentTarget) onClose();
      }}
    >
      <div
        ref={menuRef}
        className="office-file-image-context-menu"
        role="menu"
        aria-label={messages.imagePreview.contextMenu}
        style={{ left: x, top: y } as CSSProperties}
        onKeyDown={handleKeyDown}
        onPointerDown={(event) => event.stopPropagation()}
      >
        <button
          type="button"
          className="office-file-image-context-menu__item"
          role="menuitem"
          onClick={onPreview}
        >
          <PreviewIcon />
          <span>{messages.imagePreview.preview}</span>
        </button>
        {download ? (
          <button
            type="button"
            className="office-file-image-context-menu__item"
            role="menuitem"
            disabled={!resource.url || downloading}
            onClick={handleDownload}
          >
            <DownloadIcon />
            <span>{messages.imagePreview.download}</span>
          </button>
        ) : null}
        {errorMessage ? (
          <div className="office-file-image-context-menu__error" role="alert">
            {errorMessage}
          </div>
        ) : null}
      </div>
    </div>
  );
}
