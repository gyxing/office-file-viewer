// OfficeToolbar 提供打开文件、翻页、缩放、全屏等 OfficeFileViewer 顶部通用操作。
import { Button, Select, Space, Tooltip, Typography, Upload } from 'antd';
import React, { memo, useMemo } from 'react';
import { useOfficeFileViewerMessages } from '../locale';
import type { PreviewKind } from '../services/preview';
import {
  isPresentationPreviewKind,
  isWordPreviewKind,
} from '../services/preview';
import {
  OFFICE_DEFAULT_ZOOM,
  OFFICE_MAX_ZOOM,
  OFFICE_MIN_ZOOM,
  OFFICE_ZOOM_LEVELS,
} from './constants';
import { OfficeFileTypeIcon } from './FileTypeIcon';
import {
  ChevronLeftIcon,
  ChevronRightIcon,
  FolderOpenIcon,
  FullscreenIcon,
  NotesIcon,
  OutlineIcon,
  ZoomInIcon,
  ZoomOutIcon,
} from './icons';
/** 文件选择器允许选择的 Office 扩展名列表。 */
const OFFICE_FILE_ACCEPT = '.pptx,.ppt,.xlsx,.xls,.docx,.doc,.wps';

/** Office工具栏组件属性。 */
type OfficeToolbarProps = {
  /** 正在解析的原始文件名，用于格式识别和错误提示。 */
  fileName: string;
  /** 当前文件识别出的预览格式；尚未选择文件时为空。 */
  previewKind?: PreviewKind;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 当前是否已有可预览的文档。 */
  hasDocument: boolean;
  /** 当前幻灯片是否允许切换到上一页。 */
  canGoPreviousSlide: boolean;
  /** 当前幻灯片是否允许切换到下一页。 */
  canGoNextSlide: boolean;
  /** 演讲者备注面板当前是否展开。 */
  showSpeakerNotes: boolean;
  /** 切换演讲者备注面板的展开状态。 */
  onToggleSpeakerNotes: () => void;
  /** 当前文字文档是否包含可导航的大纲。 */
  hasWordOutline: boolean;
  /** 文字文档大纲当前是否展开。 */
  showWordOutline: boolean;
  /** 切换文字文档大纲的展开状态。 */
  onToggleWordOutline: () => void;
  /** 在 SelectFile 事件发生时调用的回调函数。 */
  onSelectFile: (file: File) => void;
  /** 在 PreviousSlide 事件发生时调用的回调函数。 */
  onPreviousSlide: () => void;
  /** 在 NextSlide 事件发生时调用的回调函数。 */
  onNextSlide: () => void;
  /** 在 ZoomOut 事件发生时调用的回调函数。 */
  onZoomOut: () => void;
  /** 在 ZoomIn 事件发生时调用的回调函数。 */
  onZoomIn: () => void;
  /** 在 ZoomChange 事件发生时调用的回调函数。 */
  onZoomChange: (zoom: number) => void;
  /** 在 ResetZoom 事件发生时调用的回调函数。 */
  onResetZoom: () => void;
  /** 预览器当前是否处于全屏状态。 */
  isFullscreen: boolean;
  /** 当前运行环境是否支持全屏 API。 */
  fullscreenSupported: boolean;
  /** 在 Fullscreen 事件发生时调用的回调函数。 */
  onFullscreen: () => void;
};

/** 提供缩放、翻页、备注和全屏等预览操作。 */
function OfficeToolbarComponent({
  fileName,
  previewKind,
  zoom,
  hasDocument,
  canGoPreviousSlide,
  canGoNextSlide,
  showSpeakerNotes,
  onToggleSpeakerNotes,
  hasWordOutline,
  showWordOutline,
  onToggleWordOutline,
  onSelectFile,
  onPreviousSlide,
  onNextSlide,
  onZoomOut,
  onZoomIn,
  onZoomChange,
  onResetZoom,
  isFullscreen,
  fullscreenSupported,
  onFullscreen,
}: OfficeToolbarProps) {
  const messages = useOfficeFileViewerMessages();
  const zoomOptions = useMemo(
    () => OFFICE_ZOOM_LEVELS.map((value) => ({ value, label: `${value}%` })),
    [],
  );
  const isPresentationPreview = previewKind
    ? isPresentationPreviewKind(previewKind)
    : false;
  const isWordPreview = previewKind ? isWordPreviewKind(previewKind) : false;
  // 只有已加载且至少存在一个可切换方向的演示文稿才需要显示翻页导航。
  const showSlideNavigation =
    isPresentationPreview &&
    hasDocument &&
    (canGoPreviousSlide || canGoNextSlide);
  const speakerNotesLabel = showSpeakerNotes
    ? messages.toolbar.hideSpeakerNotes
    : messages.toolbar.showSpeakerNotes;
  const showWordOutlineToggle = isWordPreview && hasDocument && hasWordOutline;
  const wordOutlineLabel = showWordOutline
    ? messages.outline.collapse
    : messages.outline.expand;
  const fullscreenLabel = isFullscreen
    ? messages.toolbar.exitFullscreen
    : messages.toolbar.fullscreen;

  return (
    <div className="office-file-toolbar">
      <div className="office-file-toolbar__file-info">
        <OfficeFileTypeIcon
          className="office-file-toolbar__filename-icon"
          previewKind={previewKind}
        />
        <Typography.Text
          strong
          ellipsis
          className="office-file-toolbar__filename"
        >
          {fileName}
        </Typography.Text>
      </div>
      <Space size={8} wrap>
        <Upload
          accept={OFFICE_FILE_ACCEPT}
          showUploadList={false}
          beforeUpload={(file) => {
            void onSelectFile(file);
            return false;
          }}
        >
          <Button icon={<FolderOpenIcon />}>
            {messages.toolbar.selectFile}
          </Button>
        </Upload>
        {showSlideNavigation ? (
          <>
            <Tooltip title={messages.toolbar.previousSlide}>
              <Button
                aria-label={messages.toolbar.previousSlide}
                icon={<ChevronLeftIcon />}
                disabled={!canGoPreviousSlide}
                onClick={onPreviousSlide}
              />
            </Tooltip>
            <Tooltip title={messages.toolbar.nextSlide}>
              <Button
                aria-label={messages.toolbar.nextSlide}
                icon={<ChevronRightIcon />}
                disabled={!canGoNextSlide}
                onClick={onNextSlide}
              />
            </Tooltip>
          </>
        ) : null}
        {isPresentationPreview ? (
          <Tooltip title={speakerNotesLabel}>
            <Button
              aria-label={speakerNotesLabel}
              aria-pressed={showSpeakerNotes}
              type={showSpeakerNotes ? 'primary' : 'default'}
              icon={<NotesIcon />}
              disabled={!hasDocument}
              onClick={onToggleSpeakerNotes}
            >
              {messages.toolbar.speakerNotes}
            </Button>
          </Tooltip>
        ) : null}
        {showWordOutlineToggle ? (
          <Tooltip title={wordOutlineLabel}>
            <Button
              aria-label={wordOutlineLabel}
              aria-pressed={showWordOutline}
              type={showWordOutline ? 'primary' : 'default'}
              icon={<OutlineIcon />}
              onClick={onToggleWordOutline}
            >
              {messages.outline.title}
            </Button>
          </Tooltip>
        ) : null}
        <Select
          value={zoom}
          className="office-file-toolbar__zoom"
          onChange={onZoomChange}
          options={zoomOptions}
        />
        <Tooltip title={messages.toolbar.zoomOut}>
          <Button
            aria-label={messages.toolbar.zoomOut}
            icon={<ZoomOutIcon />}
            disabled={!hasDocument || zoom <= OFFICE_MIN_ZOOM}
            onClick={onZoomOut}
          />
        </Tooltip>
        <Tooltip title={messages.toolbar.zoomIn}>
          <Button
            aria-label={messages.toolbar.zoomIn}
            icon={<ZoomInIcon />}
            disabled={!hasDocument || zoom >= OFFICE_MAX_ZOOM}
            onClick={onZoomIn}
          />
        </Tooltip>
        <Button disabled={!hasDocument} onClick={onResetZoom}>
          {OFFICE_DEFAULT_ZOOM}%
        </Button>
        <Button
          aria-label={fullscreenLabel}
          icon={<FullscreenIcon />}
          disabled={!hasDocument || !fullscreenSupported}
          onClick={onFullscreen}
        >
          {fullscreenLabel}
        </Button>
      </Space>
    </div>
  );
}

export const OfficeToolbar = memo(OfficeToolbarComponent);
