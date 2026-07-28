// OfficeToolbar 提供选择文件、翻页、缩放、全屏等 OfficeFileViewer 顶部通用操作。
import { Button, Select, Space, Tooltip, Typography, Upload } from 'antd';
import React, { memo, useMemo } from 'react';
import type { PreviewKind } from '../services/preview';
import {
  isPresentationPreviewKind,
  isSpreadsheetPreviewKind,
} from '../services/preview';
import {
  OFFICE_DEFAULT_ZOOM,
  OFFICE_MAX_ZOOM,
  OFFICE_MIN_ZOOM,
  OFFICE_ZOOM_LEVELS,
} from './constants';
import {
  ChevronLeftIcon,
  ChevronRightIcon,
  FileExcelIcon,
  FilePptIcon,
  FileWordIcon,
  FullscreenIcon,
  NotesIcon,
  ZoomInIcon,
  ZoomOutIcon,
} from './icons';
const OFFICE_FILE_ACCEPT = '.pptx,.ppt,.xlsx,.xls,.docx,.doc,.wps';

/** 定义 OfficeToolbar 组件可接收的属性。 */
type OfficeToolbarProps = {
  /** 正在解析的原始文件名，用于格式识别和错误提示。 */
  fileName: string;
  /** 当前文件识别出的预览格式。 */
  previewKind: PreviewKind;
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

/** 获取 `getPreviewIcon` 返回的数据。 */
function getPreviewIcon(kind: PreviewKind) {
  if (isSpreadsheetPreviewKind(kind)) return <FileExcelIcon />;
  if (kind === 'docx' || kind === 'doc') return <FileWordIcon />;
  return <FilePptIcon />;
}

/** 渲染 OfficeToolbarComponent 组件。 */
function OfficeToolbarComponent({
  fileName,
  previewKind,
  zoom,
  hasDocument,
  canGoPreviousSlide,
  canGoNextSlide,
  showSpeakerNotes,
  onToggleSpeakerNotes,
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
  const zoomOptions = useMemo(
    () => OFFICE_ZOOM_LEVELS.map((value) => ({ value, label: `${value}%` })),
    [],
  );
  // 只有已加载且至少存在一个可切换方向的演示文稿才需要显示翻页导航。
  const showSlideNavigation =
    isPresentationPreviewKind(previewKind) &&
    hasDocument &&
    (canGoPreviousSlide || canGoNextSlide);

  return (
    <div className="office-file-toolbar">
      <Typography.Text
        strong
        ellipsis
        className="office-file-toolbar__filename"
      >
        {fileName}
      </Typography.Text>
      <Space size={8} wrap>
        <Upload
          accept={OFFICE_FILE_ACCEPT}
          showUploadList={false}
          beforeUpload={(file) => {
            void onSelectFile(file);
            return false;
          }}
        >
          <Button icon={getPreviewIcon(previewKind)}>选择文件</Button>
        </Upload>
        {showSlideNavigation ? (
          <>
            <Tooltip title="上一页">
              <Button
                aria-label="上一页"
                icon={<ChevronLeftIcon />}
                disabled={!canGoPreviousSlide}
                onClick={onPreviousSlide}
              />
            </Tooltip>
            <Tooltip title="下一页">
              <Button
                aria-label="下一页"
                icon={<ChevronRightIcon />}
                disabled={!canGoNextSlide}
                onClick={onNextSlide}
              />
            </Tooltip>
          </>
        ) : null}
        {isPresentationPreviewKind(previewKind) ? (
          <Tooltip
            title={showSpeakerNotes ? '隐藏演讲者备注' : '显示演讲者备注'}
          >
            <Button
              aria-label={
                showSpeakerNotes ? '隐藏演讲者备注' : '显示演讲者备注'
              }
              aria-pressed={showSpeakerNotes}
              type={showSpeakerNotes ? 'primary' : 'default'}
              icon={<NotesIcon />}
              disabled={!hasDocument}
              onClick={onToggleSpeakerNotes}
            >
              备注
            </Button>
          </Tooltip>
        ) : null}
        <Select
          value={zoom}
          className="office-file-toolbar__zoom"
          onChange={onZoomChange}
          options={zoomOptions}
        />
        <Tooltip title="缩小">
          <Button
            aria-label="缩小"
            icon={<ZoomOutIcon />}
            disabled={!hasDocument || zoom <= OFFICE_MIN_ZOOM}
            onClick={onZoomOut}
          />
        </Tooltip>
        <Tooltip title="放大">
          <Button
            aria-label="放大"
            icon={<ZoomInIcon />}
            disabled={!hasDocument || zoom >= OFFICE_MAX_ZOOM}
            onClick={onZoomIn}
          />
        </Tooltip>
        <Button disabled={!hasDocument} onClick={onResetZoom}>
          {OFFICE_DEFAULT_ZOOM}%
        </Button>
        <Button
          aria-label={isFullscreen ? '退出全屏' : '全屏'}
          icon={<FullscreenIcon />}
          disabled={!hasDocument || !fullscreenSupported}
          onClick={onFullscreen}
        >
          {isFullscreen ? '退出全屏' : '全屏'}
        </Button>
      </Space>
    </div>
  );
}

export const OfficeToolbar = memo(OfficeToolbarComponent);
