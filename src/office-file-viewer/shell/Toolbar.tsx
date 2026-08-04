// OfficeToolbar 提供打开文件、翻页、缩放、全屏等 OfficeFileViewer 顶部通用操作。
import { Button, Select, Space, Tooltip, Typography, Upload } from 'antd';
import React, { memo, useMemo } from 'react';
import {
  useOfficeFileViewerMessages,
  type OfficeFileViewerMessages,
} from '../locale';
import type { PreviewKind } from '../services/preview';
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

/** 工具栏按文件格式开放的互斥操作能力。 */
export type OfficeToolbarFormatControls =
  | { kind: 'empty' }
  | {
      kind: 'presentation';
      /** 多页演示文稿的翻页能力；单页时不提供。 */
      navigation?: {
        /** 当前页是否允许向前切换。 */
        canPrevious: boolean;
        /** 当前页是否允许向后切换。 */
        canNext: boolean;
        /** 切换到上一张幻灯片。 */
        previous(): void;
        /** 切换到下一张幻灯片。 */
        next(): void;
      };
      /** 演讲者备注面板的受控操作能力。 */
      speakerNotes: {
        /** 备注面板当前是否展开。 */
        visible: boolean;
        /** 当前是否尚无可展示的演示文稿内容。 */
        disabled: boolean;
        /** 切换备注面板的展开状态。 */
        toggle(): void;
      };
    }
  | {
      kind: 'word';
      /** 文档存在真实大纲时提供的显隐能力。 */
      outline?: {
        /** 大纲侧栏当前是否展开。 */
        visible: boolean;
        /** 切换大纲侧栏的展开状态。 */
        toggle(): void;
      };
    }
  | { kind: 'spreadsheet' };

/** 工具栏通用缩放能力。 */
export type ZoomControls = {
  /** 当前预览缩放比例。 */
  value: number;
  /** 当前是否已有可执行增减和复位操作的文档。 */
  hasDocument: boolean;
  /** 减小预览比例。 */
  zoomOut(): void;
  /** 放大预览比例。 */
  zoomIn(): void;
  /** 设置指定预览比例。 */
  change(value: number): void;
  /** 恢复默认预览比例。 */
  reset(): void;
};

/** 工具栏通用全屏能力。 */
export type FullscreenControls = {
  /** 预览器当前是否处于全屏状态。 */
  active: boolean;
  /** 当前是否因无文档或浏览器不支持而不可用。 */
  disabled: boolean;
  /** 进入或退出浏览器全屏。 */
  toggle(): void;
};

/** 顶部工具栏组合所需的文件与操作能力。 */
type OfficeToolbarProps = {
  /** 当前显示的文件名。 */
  fileName: string;
  /** 当前文件识别出的精确预览格式。 */
  previewKind?: PreviewKind;
  /** 当前格式专属的互斥操作能力。 */
  formatControls: OfficeToolbarFormatControls;
  /** 通用缩放能力。 */
  zoomControls: ZoomControls;
  /** 通用全屏能力。 */
  fullscreenControls: FullscreenControls;
  /** 解析用户在文件选择器中选中的文件。 */
  onSelectFile(file: File): void;
};

type PresentationControls = Extract<
  OfficeToolbarFormatControls,
  { kind: 'presentation' }
>;

type WordControls = Extract<OfficeToolbarFormatControls, { kind: 'word' }>;

/** 演示文稿专属操作块所需的数据。 */
type PresentationControlGroupProps = {
  /** 翻页和备注能力。 */
  controls: PresentationControls;
  /** 当前语言环境对应的界面文案。 */
  messages: OfficeFileViewerMessages;
};

/** Word 大纲操作块所需的数据。 */
type WordControlProps = {
  /** 当前可用的大纲能力。 */
  controls: WordControls;
  /** 当前语言环境对应的界面文案。 */
  messages: OfficeFileViewerMessages;
};

/** 缩放操作块所需的数据。 */
type ZoomControlGroupProps = {
  /** 当前可用的缩放能力。 */
  controls: ZoomControls;
  /** 当前语言环境对应的界面文案。 */
  messages: OfficeFileViewerMessages;
  /** 可选缩放比例。 */
  options: Array<{ value: number; label: string }>;
};

/** 全屏操作块所需的数据。 */
type FullscreenControlProps = {
  /** 当前可用的全屏能力。 */
  controls: FullscreenControls;
  /** 当前语言环境对应的界面文案。 */
  messages: OfficeFileViewerMessages;
};

/** 按 Space 的直接子项结构渲染演示文稿操作。 */
function renderPresentationControls({
  controls,
  messages,
}: PresentationControlGroupProps) {
  const speakerNotesLabel = controls.speakerNotes.visible
    ? messages.toolbar.hideSpeakerNotes
    : messages.toolbar.showSpeakerNotes;
  const items = [];

  if (controls.navigation) {
    items.push(
      <Tooltip key="previous" title={messages.toolbar.previousSlide}>
        <Button
          aria-label={messages.toolbar.previousSlide}
          icon={<ChevronLeftIcon />}
          disabled={!controls.navigation.canPrevious}
          onClick={controls.navigation.previous}
        />
      </Tooltip>,
      <Tooltip key="next" title={messages.toolbar.nextSlide}>
        <Button
          aria-label={messages.toolbar.nextSlide}
          icon={<ChevronRightIcon />}
          disabled={!controls.navigation.canNext}
          onClick={controls.navigation.next}
        />
      </Tooltip>,
    );
  }
  items.push(
    <Tooltip key="speaker-notes" title={speakerNotesLabel}>
      <Button
        aria-label={speakerNotesLabel}
        aria-pressed={controls.speakerNotes.visible}
        type={controls.speakerNotes.visible ? 'primary' : 'default'}
        icon={<NotesIcon />}
        disabled={controls.speakerNotes.disabled}
        onClick={controls.speakerNotes.toggle}
      >
        {messages.toolbar.speakerNotes}
      </Button>
    </Tooltip>,
  );
  return items;
}

/** 仅在文档具有真实大纲时渲染显隐入口。 */
function WordControl({ controls, messages }: WordControlProps) {
  if (!controls.outline) return null;
  const outlineLabel = controls.outline.visible
    ? messages.outline.collapse
    : messages.outline.expand;

  return (
    <Tooltip title={outlineLabel}>
      <Button
        aria-label={outlineLabel}
        aria-pressed={controls.outline.visible}
        type={controls.outline.visible ? 'primary' : 'default'}
        icon={<OutlineIcon />}
        onClick={controls.outline.toggle}
      >
        {messages.outline.title}
      </Button>
    </Tooltip>
  );
}

/** 按 Space 的直接子项结构渲染通用缩放操作。 */
function renderZoomControls({
  controls,
  messages,
  options,
}: ZoomControlGroupProps) {
  return [
    <Select
      key="zoom-select"
      value={controls.value}
      className="office-file-toolbar__zoom"
      onChange={controls.change}
      options={options}
    />,
    <Tooltip key="zoom-out" title={messages.toolbar.zoomOut}>
      <Button
        aria-label={messages.toolbar.zoomOut}
        icon={<ZoomOutIcon />}
        disabled={!controls.hasDocument || controls.value <= OFFICE_MIN_ZOOM}
        onClick={controls.zoomOut}
      />
    </Tooltip>,
    <Tooltip key="zoom-in" title={messages.toolbar.zoomIn}>
      <Button
        aria-label={messages.toolbar.zoomIn}
        icon={<ZoomInIcon />}
        disabled={!controls.hasDocument || controls.value >= OFFICE_MAX_ZOOM}
        onClick={controls.zoomIn}
      />
    </Tooltip>,
    <Button
      key="zoom-reset"
      disabled={!controls.hasDocument}
      onClick={controls.reset}
    >
      {OFFICE_DEFAULT_ZOOM}%
    </Button>,
  ];
}

/** 渲染所有格式共享的全屏操作。 */
function FullscreenControl({ controls, messages }: FullscreenControlProps) {
  const label = controls.active
    ? messages.toolbar.exitFullscreen
    : messages.toolbar.fullscreen;

  return (
    <Button
      aria-label={label}
      icon={<FullscreenIcon />}
      disabled={controls.disabled}
      onClick={controls.toggle}
    >
      {label}
    </Button>
  );
}

/** 提供打开文件、格式专属操作、缩放和全屏等预览能力。 */
function OfficeToolbarComponent({
  fileName,
  previewKind,
  formatControls,
  zoomControls,
  fullscreenControls,
  onSelectFile,
}: OfficeToolbarProps) {
  const messages = useOfficeFileViewerMessages();
  const zoomOptions = useMemo(() => {
    // 缩放按钮会产生非预设档位，补入当前值后 Select 才能持续显示百分号标签。
    const levels = OFFICE_ZOOM_LEVELS.includes(zoomControls.value)
      ? OFFICE_ZOOM_LEVELS
      : [...OFFICE_ZOOM_LEVELS, zoomControls.value].sort(
          (left, right) => left - right,
        );
    return levels.map((value) => ({ value, label: `${value}%` }));
  }, [zoomControls.value]);

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
        {formatControls.kind === 'presentation'
          ? renderPresentationControls({ controls: formatControls, messages })
          : null}
        {formatControls.kind === 'word' ? (
          <WordControl controls={formatControls} messages={messages} />
        ) : null}
        {renderZoomControls({
          controls: zoomControls,
          messages,
          options: zoomOptions,
        })}
        <FullscreenControl controls={fullscreenControls} messages={messages} />
      </Space>
    </div>
  );
}

export const OfficeToolbar = memo(OfficeToolbarComponent);
