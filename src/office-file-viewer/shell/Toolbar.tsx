// OfficeToolbar 提供打开文件、翻页、缩放、全屏等 OfficeFileViewer 顶部通用操作。
import type { ChangeEvent, ReactElement, ReactNode } from 'react';
import React, { memo, useRef } from 'react';
import {
  useOfficeFileViewerMessages,
  type OfficeFileViewerMessages,
} from '../locale';
import type { WordRevisionMode } from '../services/annotations/types';
import { OFFICE_FILE_ACCEPT } from '../services/parsing/formatDefinitions';
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
  ReviewIcon,
  SearchIcon,
} from '../shared/ui/OfficeIcons';
import { OfficeTooltip } from '../shared/ui/OfficeTooltip';
import { OfficeFileTypeIcon } from './OfficeFileTypeIcon';
import { SpreadsheetViewModeControl } from './SpreadsheetViewModeControl';
import { WordRevisionModeControl } from './WordRevisionModeControl';
import { ZoomControl, type ZoomControls } from './ZoomControl';

export type { ZoomControls } from './ZoomControl';
/** 宿主可以独立隐藏的内置工具栏区域。 */
export type OfficeFileViewerToolbarOptions = Readonly<{
  /** 是否显示文件名和格式图标，默认显示。 */
  fileName?: boolean;
  /** 是否显示内置文件选择入口，默认显示。 */
  openFile?: boolean;
  /** 是否显示 Word 文档大纲入口，默认显示。 */
  wordOutline?: boolean;
  /** 是否显示 Word 修订投影切换，默认显示。 */
  wordRevisionMode?: boolean;
  /** 是否显示电子表格显示模式切换，默认显示。 */
  spreadsheetViewMode?: boolean;
  /** 是否显示演示文稿上一页和下一页操作，默认显示。 */
  presentationNavigation?: boolean;
  /** 是否显示演讲者备注入口，默认显示。 */
  speakerNotes?: boolean;
  /** 是否显示全文查找入口，默认显示。 */
  search?: boolean;
  /** 是否显示文档审阅入口，默认显示。 */
  review?: boolean;
  /** 是否显示缩放控件，默认显示。 */
  zoom?: boolean;
  /** 是否显示全屏操作，默认显示。 */
  fullscreen?: boolean;
}>;

/** 工具栏实际渲染使用的完整显示配置。 */
export type ResolvedOfficeFileViewerToolbarOptions =
  Required<OfficeFileViewerToolbarOptions>;
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
      /** 文档包含修订时提供的只读投影切换能力。 */
      revisionMode?: {
        /** 当前采用的修订投影。 */
        value: WordRevisionMode;
        /** 当前是否禁止切换修订投影。 */
        disabled: boolean;
        /** 切换修订投影。 */
        change(value: WordRevisionMode): void;
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

/** 工具栏文档审阅入口的显式能力。 */
export type OfficeToolbarReviewControls =
  | { kind: 'disabled' }
  | {
      kind: 'enabled';
      /** 审阅侧栏当前是否展开。 */
      visible: boolean;
      /** 当前文档是否尚无可展示的审阅内容。 */
      disabled: boolean;
      /** 当前批注、修订和笔记的合计数量。 */
      count: number;
      /** 切换审阅侧栏。 */
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
  /** 文档审阅入口能力。 */
  reviewControls: OfficeToolbarReviewControls;
  /** 内置操作区域的完整显示配置。 */
  displayOptions: ResolvedOfficeFileViewerToolbarOptions;
  /** 宿主追加到全部内置操作之后的自定义工具栏内容。 */
  extra?: ReactNode;
  /** 文件选择器接受的文件类型，默认使用全部受支持的 Office 格式。 */
  fileAccept?: string;
  /** 解析用户在文件选择器中选中的文件；为空时不显示打开文件入口。 */
  onSelectFile?: (file: File) => void;
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
  /** 演示文稿专属入口的显示配置。 */
  displayOptions: ResolvedOfficeFileViewerToolbarOptions;
  /** 当前语言环境对应的界面文案。 */
  messages: OfficeFileViewerMessages;
};

/** Word 大纲操作块所需的数据。 */
type WordControlProps = {
  /** 当前可用的大纲能力。 */
  controls: WordControls;
  /** Word 专属入口的显示配置。 */
  displayOptions: ResolvedOfficeFileViewerToolbarOptions;
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
  displayOptions,
  messages,
}: PresentationControlGroupProps) {
  const speakerNotesLabel = controls.speakerNotes.visible
    ? messages.toolbar.hideSpeakerNotes
    : messages.toolbar.showSpeakerNotes;
  const items = [];

  if (displayOptions.presentationNavigation && controls.navigation) {
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
  if (displayOptions.speakerNotes) {
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
  }
  return items;
}

/** 渲染 Word 大纲入口和修订投影切换。 */
function WordControl({ controls, displayOptions, messages }: WordControlProps) {
  if (!controls.outline && !controls.revisionMode) return null;
  const outlineLabel = controls.outline
    ? controls.outline.visible
      ? messages.outline.collapse
      : messages.outline.expand
    : undefined;

  return (
    <>
      {displayOptions.wordOutline && controls.outline ? (
        <OfficeButton
          aria-label={outlineLabel}
          aria-pressed={controls.outline.visible}
          variant={controls.outline.visible ? 'primary' : 'default'}
          icon={<OutlineIcon />}
          onClick={controls.outline.toggle}
        >
          {messages.outline.title}
        </OfficeButton>
      ) : null}
      {displayOptions.wordRevisionMode && controls.revisionMode ? (
        <WordRevisionModeControl
          value={controls.revisionMode.value}
          disabled={controls.revisionMode.disabled}
          onChange={controls.revisionMode.change}
        />
      ) : null}
    </>
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
  reviewControls,
  displayOptions,
  extra,
  fileAccept = OFFICE_FILE_ACCEPT,
  onSelectFile,
}: OfficeToolbarProps): ReactElement {
  const messages = useOfficeFileViewerMessages();
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.currentTarget.files?.[0];
    // 清空原生输入值，确保用户可以连续打开同一文件。
    event.currentTarget.value = '';
    if (file) void onSelectFile?.(file);
  };

  return (
    <div
      className="office-file-toolbar"
      role="toolbar"
      aria-label={messages.toolbar.region}
    >
      {displayOptions.fileName ? (
        <div className="office-file-toolbar__file-info">
          <OfficeFileTypeIcon
            className="office-file-toolbar__filename-icon"
            previewKind={previewKind}
          />
          <strong className="office-file-toolbar__filename" title={fileName}>
            {fileName}
          </strong>
        </div>
      ) : null}
      <div className="office-file-toolbar__actions">
        {displayOptions.openFile && onSelectFile ? (
          <>
            <input
              ref={fileInputRef}
              className="office-file-toolbar__file-input"
              type="file"
              accept={fileAccept}
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
          </>
        ) : null}
        {formatControls.kind === 'presentation'
          ? renderPresentationControls({
              controls: formatControls,
              displayOptions,
              messages,
            })
          : null}
        {formatControls.kind === 'word' ? (
          <WordControl
            controls={formatControls}
            displayOptions={displayOptions}
            messages={messages}
          />
        ) : null}
        {displayOptions.spreadsheetViewMode &&
        formatControls.kind === 'spreadsheet' ? (
          <SpreadsheetViewModeControl
            value={formatControls.viewMode.value}
            disabled={formatControls.viewMode.disabled}
            onChange={formatControls.viewMode.change}
          />
        ) : null}
        {displayOptions.search && searchControls.kind === 'enabled' ? (
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
        {displayOptions.review && reviewControls.kind === 'enabled' ? (
          <OfficeButton
            data-testid="office-review-toggle"
            aria-label={
              reviewControls.visible
                ? messages.review.collapse
                : messages.review.expand
            }
            aria-pressed={reviewControls.visible}
            variant={reviewControls.visible ? 'primary' : 'default'}
            icon={<ReviewIcon />}
            disabled={reviewControls.disabled}
            onClick={reviewControls.toggle}
          >
            {messages.review.title}
          </OfficeButton>
        ) : null}
        {displayOptions.zoom ? <ZoomControl controls={zoomControls} /> : null}
        {displayOptions.fullscreen ? (
          <FullscreenControl
            controls={fullscreenControls}
            messages={messages}
          />
        ) : null}
        {extra ? (
          <div className="office-file-toolbar__extra">{extra}</div>
        ) : null}
      </div>
    </div>
  );
}

export const OfficeToolbar = memo(OfficeToolbarComponent);
