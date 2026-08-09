import type { ImgHTMLAttributes, KeyboardEvent, MouseEvent } from 'react';
import React, { memo, useMemo } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type { OfficeResourceSource } from '../../services/resource-store';
import { useOfficeImagePreviewContext } from './OfficeImagePreviewContext';
import type { OfficeImagePreviewTarget } from './types';

/** 可预览图片在原生 img 属性之外需要的资源元数据。 */
export type OfficePreviewableImageProps = Omit<
  ImgHTMLAttributes<HTMLImageElement>,
  'src'
> & {
  /** 当前 img 元素实际显示的浏览器资源地址。 */
  src?: string;
  /** 预览层需要独立持有的原始图片资源。 */
  previewSource: string | OfficeResourceSource;
  /** 当前图片在所属文档内的稳定标识。 */
  previewId: string;
  /** 优先用于预览标题和下载文件名的图片名称。 */
  previewName?: string;
  /** 用于推断下载扩展名的图片 MIME 类型。 */
  previewMimeType?: string;
  /** 当前图片是否同时承载文档超链接。 */
  'data-office-hyperlink'?: 'true';
};

/** 为可交互内容图片补充双击、右键菜单和键盘入口。 */
function OfficePreviewableImageComponent({
  alt,
  className,
  onContextMenu,
  onDoubleClick,
  onKeyDown,
  previewId,
  previewMimeType,
  previewName,
  previewSource,
  role,
  src,
  tabIndex,
  'data-office-hyperlink': hyperlinkData,
  ...imageProps
}: OfficePreviewableImageProps) {
  const messages = useOfficeFileViewerMessages();
  const previewContext = useOfficeImagePreviewContext();
  const target = useMemo<OfficeImagePreviewTarget>(
    () => ({
      id: previewId,
      source: previewSource,
      alt,
      name: previewName,
      mimeType: previewMimeType,
    }),
    [alt, previewId, previewMimeType, previewName, previewSource],
  );
  const interactive = Boolean(previewContext?.options.enabled && src);
  const hasHyperlink = hyperlinkData === 'true';
  const openPreview = (trigger: HTMLImageElement) => {
    if (interactive) previewContext?.openPreview(target, trigger);
  };
  const openContextMenu = (
    trigger: HTMLImageElement,
    clientX: number,
    clientY: number,
  ) => {
    if (!interactive || !previewContext?.options.contextMenu) return;
    previewContext.openContextMenu(target, { clientX, clientY }, trigger);
  };
  const handleDoubleClick = (event: MouseEvent<HTMLImageElement>) => {
    onDoubleClick?.(event);
    // 修饰键单击属于链接事务，双击时不能再叠加打开图片预览。
    if (!event.defaultPrevented && !event.ctrlKey && !event.metaKey) {
      openPreview(event.currentTarget);
    }
  };
  const handleContextMenu = (event: MouseEvent<HTMLImageElement>) => {
    onContextMenu?.(event);
    if (
      event.defaultPrevented ||
      !interactive ||
      !previewContext?.options.contextMenu
    ) {
      return;
    }
    event.preventDefault();
    event.currentTarget.focus();
    openContextMenu(event.currentTarget, event.clientX, event.clientY);
  };
  const handleKeyDown = (event: KeyboardEvent<HTMLImageElement>) => {
    onKeyDown?.(event);
    if (event.defaultPrevented || !interactive) return;
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      openPreview(event.currentTarget);
      return;
    }
    if (
      event.key === 'ContextMenu' ||
      (event.shiftKey && event.key === 'F10')
    ) {
      if (!previewContext?.options.contextMenu) return;
      event.preventDefault();
      const rect = event.currentTarget.getBoundingClientRect();
      openContextMenu(
        event.currentTarget,
        rect.left + Math.min(rect.width, 24),
        rect.top + Math.min(rect.height, 24),
      );
    }
  };
  const label = previewName || alt || '';

  return (
    <img
      {...imageProps}
      src={src}
      alt={alt}
      className={
        [className, interactive ? 'office-file-previewable-image' : undefined]
          .filter(Boolean)
          .join(' ') || undefined
      }
      role={hasHyperlink ? role ?? 'link' : interactive ? 'button' : role}
      tabIndex={hasHyperlink || interactive ? tabIndex ?? 0 : tabIndex}
      aria-label={
        hasHyperlink
          ? imageProps['aria-label'] ?? alt
          : interactive
          ? messages.imagePreview.openLabel(label)
          : imageProps['aria-label']
      }
      data-office-hyperlink={hyperlinkData}
      data-office-image-previewable={interactive ? 'true' : undefined}
      onContextMenu={handleContextMenu}
      onDoubleClick={handleDoubleClick}
      onKeyDown={handleKeyDown}
    />
  );
}

export const OfficePreviewableImage = memo(OfficePreviewableImageComponent);
