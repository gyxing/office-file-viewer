import type { KeyboardEvent, MouseEvent } from 'react';
import React, { useEffect, useRef, useState } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import { OfficeButton } from '../ui/OfficeButton';
import {
  CloseIcon,
  DownloadIcon,
  ResetIcon,
  RotateClockwiseIcon,
  ZoomInIcon,
  ZoomOutIcon,
} from '../ui/OfficeIcons';
import { OfficeSpinner } from '../ui/OfficeSpinner';
import { downloadOfficeImage } from './downloadOfficeImage';
import type { OfficeImagePreviewTarget } from './types';
import { useOfficeImagePreviewResource } from './useOfficeImagePreviewResource';
import { useOfficeImagePreviewTransform } from './useOfficeImagePreviewTransform';

/** 图片预览弹层需要的资源与关闭能力。 */
type OfficeImagePreviewDialogProps = {
  /** 当前需要查看的图片。 */
  target: OfficeImagePreviewTarget;
  /** 是否显示下载原始图片操作。 */
  download: boolean;
  /** 当前预览器所在的文档对象。 */
  ownerDocument: Document;
  /** 关闭图片预览层。 */
  onClose(): void;
};

/** 获取弹层内当前可通过 Tab 到达的控件。 */
function getDialogFocusableElements(dialog: HTMLElement) {
  return Array.from(
    dialog.querySelectorAll<HTMLElement>(
      'button:not(:disabled), [href], [tabindex]:not([tabindex="-1"])',
    ),
  );
}

/** 渲染支持适应窗口、缩放、拖拽、旋转和下载的图片预览层。 */
export function OfficeImagePreviewDialog({
  target,
  download,
  ownerDocument,
  onClose,
}: OfficeImagePreviewDialogProps) {
  const messages = useOfficeFileViewerMessages();
  const dialogRef = useRef<HTMLDivElement>(null);
  const [resourceGeneration, setResourceGeneration] = useState(0);
  const [imageFailed, setImageFailed] = useState(false);
  const [imageLoaded, setImageLoaded] = useState(false);
  const [downloading, setDownloading] = useState(false);
  const [downloadError, setDownloadError] = useState<string>();
  const resource = useOfficeImagePreviewResource(target, resourceGeneration);
  const transform = useOfficeImagePreviewTransform(target.id);

  useEffect(() => {
    dialogRef.current
      ?.querySelector<HTMLButtonElement>('[data-image-preview-close]')
      ?.focus();
  }, []);

  const handleDialogKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      onClose();
      return;
    }
    if (event.key !== 'Tab' || !dialogRef.current) return;
    const focusable = getDialogFocusableElements(dialogRef.current);
    if (!focusable.length) return;
    const first = focusable[0];
    const last = focusable[focusable.length - 1];
    if (event.shiftKey && ownerDocument.activeElement === first) {
      event.preventDefault();
      last.focus();
    } else if (!event.shiftKey && ownerDocument.activeElement === last) {
      event.preventDefault();
      first.focus();
    }
  };
  const handleDownload = async () => {
    if (!resource.url || downloading) return;
    setDownloading(true);
    setDownloadError(undefined);
    try {
      await downloadOfficeImage(resource.url, target, ownerDocument);
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      setDownloadError(messages.imagePreview.downloadFailed(reason));
    } finally {
      setDownloading(false);
    }
  };
  const retry = () => {
    setImageFailed(false);
    setImageLoaded(false);
    setDownloadError(undefined);
    setResourceGeneration((current) => current + 1);
  };
  const handleViewportClick = (event: MouseEvent<HTMLDivElement>) => {
    // 仅由空白预览区关闭弹层，避免图片拖拽和状态控件产生误触。
    if (event.target === event.currentTarget) onClose();
  };
  const loadFailed = Boolean(resource.error || imageFailed);
  const imageLoading =
    resource.loading || Boolean(resource.url && !imageLoaded && !loadFailed);
  const title = target.name || target.alt || messages.imagePreview.region;

  return (
    <div className="office-file-image-preview-layer">
      <div
        ref={dialogRef}
        className="office-file-image-preview-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={messages.imagePreview.region}
        onKeyDown={handleDialogKeyDown}
      >
        <header className="office-file-image-preview-dialog__header">
          <strong className="office-file-image-preview-dialog__title">
            {title}
          </strong>
          <div
            className="office-file-image-preview-dialog__toolbar"
            role="toolbar"
            aria-label={messages.imagePreview.region}
          >
            <OfficeButton
              icon={<ZoomOutIcon />}
              aria-label={messages.imagePreview.zoomOut}
              title={messages.imagePreview.zoomOut}
              disabled={!transform.canZoomOut || loadFailed}
              onClick={transform.zoomOut}
            />
            <output
              className="office-file-image-preview-dialog__zoom"
              aria-label={messages.imagePreview.zoomLevel}
            >
              {transform.zoomPercent}%
            </output>
            <OfficeButton
              icon={<ZoomInIcon />}
              aria-label={messages.imagePreview.zoomIn}
              title={messages.imagePreview.zoomIn}
              disabled={!transform.canZoomIn || loadFailed}
              onClick={transform.zoomIn}
            />
            <OfficeButton
              icon={<RotateClockwiseIcon />}
              disabled={loadFailed}
              onClick={transform.rotateClockwise}
            >
              {messages.imagePreview.rotate}
            </OfficeButton>
            <OfficeButton
              icon={<ResetIcon />}
              disabled={loadFailed}
              onClick={transform.reset}
            >
              {messages.imagePreview.reset}
            </OfficeButton>{' '}
            {download ? (
              <OfficeButton
                icon={<DownloadIcon />}
                disabled={!resource.url || downloading || loadFailed}
                onClick={handleDownload}
              >
                {messages.imagePreview.download}
              </OfficeButton>
            ) : null}
            <OfficeButton
              data-image-preview-close="true"
              icon={<CloseIcon />}
              aria-label={messages.imagePreview.close}
              title={messages.imagePreview.close}
              onClick={onClose}
            />
          </div>
        </header>
        <div
          ref={transform.viewportRef}
          className="office-file-image-preview-dialog__viewport"
          data-dragging={transform.dragging ? 'true' : 'false'}
          onClick={handleViewportClick}
          onPointerCancel={transform.handlePointerCancel}
          onPointerDown={transform.handlePointerDown}
          onPointerMove={transform.handlePointerMove}
          onPointerUp={transform.handlePointerUp}
          onWheel={transform.handleWheel}
        >
          {imageLoading ? (
            <div className="office-file-image-preview-dialog__status">
              <OfficeSpinner
                size="large"
                label={messages.imagePreview.loading}
              />
            </div>
          ) : null}
          {loadFailed ? (
            <div
              className="office-file-image-preview-dialog__status"
              role="alert"
            >
              <span>{messages.imagePreview.loadFailed}</span>
              <OfficeButton onClick={retry}>
                {messages.imagePreview.retry}
              </OfficeButton>
            </div>
          ) : null}
          {resource.url && !loadFailed ? (
            <img
              key={resourceGeneration}
              ref={transform.imageRef}
              className="office-file-image-preview-dialog__image"
              src={resource.url}
              alt={target.alt ?? target.name ?? ''}
              draggable={false}
              style={transform.imageStyle}
              onDragStart={(event) => event.preventDefault()}
              onError={() => {
                setImageLoaded(false);
                setImageFailed(true);
              }}
              onLoad={(event) => {
                setImageLoaded(true);
                transform.handleImageLoad(event);
              }}
            />
          ) : null}
        </div>
        {downloadError ? (
          <div className="office-file-image-preview-dialog__error" role="alert">
            {downloadError}
          </div>
        ) : null}
      </div>
    </div>
  );
}
