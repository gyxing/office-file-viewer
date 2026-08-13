// OfficeToolbar 提供打开文件、翻页、缩放、全屏等 OfficeFileViewer 顶部通用操作。
import type { ChangeEvent } from 'react';
import React, { memo, useRef } from 'react';
import {
  useOfficeFileViewerMessages,
  type OfficeFileViewerMessages,
} from '../locale';
import type { PreviewKind } from '../services/preview';
import type { SpreadsheetViewMode } from '../services/spreadsheet/viewMode';
import { OfficeButton } from '../shared/ui/OfficeButton';
import {
  ChevronLeftIcon,
  ChevronRightIcon,
  FolderOpenIcon,
  FullscreenIcon,
  NotesIcon,
  OutlineIcon,
  SearchIcon,
} from '../shared/ui/OfficeIcons';
import { OfficeTooltip } from '../shared/ui/OfficeTooltip';
import { OfficeFileTypeIcon } from './OfficeFileTypeIcon';
import { SpreadsheetViewModeControl } from './SpreadsheetViewModeControl';
import { ZoomControl, type ZoomControls } from './ZoomControl';

export type { ZoomControls } from './ZoomControl';
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
  | {
      kind: 'spreadsheet';
      /** 电子表格显示模式的受控能力。 */
      viewMode: {
        /** 当前采用的显示模式。 */
        value: SpreadsheetViewMode;
        /** 当前是否禁止切换。 */
        disabled: boolean;
        /** 切换显示模式。 */
        change(value: SpreadsheetViewMode): void;
      };
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

/** 工具栏文档查找入口的显式能力。 */
export type OfficeToolbarSearchControls =
  | { kind: 'disabled' }
  | {
      kind: 'enabled';
      /** 查找侧栏当前是否展开。 */
      visible: boolean;
      /** 当前是否尚无可搜索内容。 */
      disabled: boolean;
      /** 切换查找侧栏。 */
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
  /** 文档查找入口能力。 */
  searchControls: OfficeToolbarSearchControls;
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
      <OfficeTooltip key="previous" content={messages.toolbar.previousSlide}>
        <OfficeButton
          aria-label={messages.toolbar.previousSlide}
          icon={<ChevronLeftIcon />}
          disabled={!controls.navigation.canPrevious}
          onClick={controls.navigation.previous}
        />
      </OfficeTooltip>,
      <OfficeTooltip key="next" content={messages.toolbar.nextSlide}>
        <OfficeButton
          aria-label={messages.toolbar.nextSlide}
          icon={<ChevronRightIcon />}
          disabled={!controls.navigation.canNext}
          onClick={controls.navigation.next}
        />
      </OfficeTooltip>,
    );
  }
  items.push(
    <OfficeButton
      key="speaker-notes"
      aria-label={speakerNotesLabel}
      aria-pressed={controls.speakerNotes.visible}
      variant={controls.speakerNotes.visible ? 'primary' : 'default'}
      icon={<NotesIcon />}
      disabled={controls.speakerNotes.disabled}
      onClick={controls.speakerNotes.toggle}
    >
      {messages.toolbar.speakerNotes}
    </OfficeButton>,
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
    <OfficeButton
      aria-label={outlineLabel}
      aria-pressed={controls.outline.visible}
      variant={controls.outline.visible ? 'primary' : 'default'}
      icon={<OutlineIcon />}
      onClick={controls.outline.toggle}
    >
      {messages.outline.title}
    </OfficeButton>
  );
}

/** 渲染所有格式共享的全屏操作。 */
function FullscreenControl({ controls, messages }: FullscreenControlProps) {
  const label = controls.active
    ? messages.toolbar.exitFullscreen
    : messages.toolbar.fullscreen;

  return (
    <OfficeButton
      aria-label={label}
      icon={<FullscreenIcon />}
      disabled={controls.disabled}
      onClick={controls.toggle}
    >
      {label}
    </OfficeButton>
  );
}

/** 提供打开文件、格式专属操作、缩放和全屏等预览能力。 */
function OfficeToolbarComponent({
  fileName,
  previewKind,
  formatControls,
  zoomControls,
  fullscreenControls,
  searchControls,
  onSelectFile,
}: OfficeToolbarProps) {
  const messages = useOfficeFileViewerMessages();
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.currentTarget.files?.[0];
    // 清空原生输入值，确保用户可以连续打开同一文件。
    event.currentTarget.value = '';
    if (file) void onSelectFile(file);
  };

  return (
    <div
      className="office-file-toolbar"
      role="toolbar"
      aria-label={messages.toolbar.region}
    >
      <div className="office-file-toolbar__file-info">
        <OfficeFileTypeIcon
          className="office-file-toolbar__filename-icon"
          previewKind={previewKind}
        />
        <strong className="office-file-toolbar__filename" title={fileName}>
          {fileName}
        </strong>
      </div>
      <div className="office-file-toolbar__actions">
        <input
          ref={fileInputRef}
          className="office-file-toolbar__file-input"
          type="file"
          accept={OFFICE_FILE_ACCEPT}
          tabIndex={-1}
          onChange={handleFileChange}
        />
        <OfficeButton
          aria-label={messages.toolbar.selectFile}
          icon={<FolderOpenIcon />}
          onClick={() => fileInputRef.current?.click()}
        >
          {messages.toolbar.selectFile}
        </OfficeButton>
        {formatControls.kind === 'presentation'
          ? renderPresentationControls({ controls: formatControls, messages })
          : null}
        {formatControls.kind === 'word' ? (
          <WordControl controls={formatControls} messages={messages} />
        ) : null}
        {formatControls.kind === 'spreadsheet' ? (
          <SpreadsheetViewModeControl
            value={formatControls.viewMode.value}
            disabled={formatControls.viewMode.disabled}
            onChange={formatControls.viewMode.change}
          />
        ) : null}
        {searchControls.kind === 'enabled' ? (
          <OfficeButton
            data-testid="office-search-toggle"
            aria-label={
              searchControls.visible
                ? messages.search.collapse
                : messages.search.expand
            }
            aria-pressed={searchControls.visible}
            variant={searchControls.visible ? 'primary' : 'default'}
            icon={<SearchIcon />}
            disabled={searchControls.disabled}
            onClick={searchControls.toggle}
          >
            {messages.search.title}
          </OfficeButton>
        ) : null}
        <ZoomControl controls={zoomControls} />
        <FullscreenControl controls={fullscreenControls} messages={messages} />
      </div>
    </div>
  );
}

export const OfficeToolbar = memo(OfficeToolbarComponent);
