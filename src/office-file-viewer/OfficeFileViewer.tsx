// OfficeFileViewer 是组件库对外主入口，负责组合本地化、控制器、工具栏和格式预览舞台。
import { Layout } from 'antd';
import type { CSSProperties, ReactElement } from 'react';
import React, { memo, useMemo } from 'react';
import './index.less';
import {
  OfficeFileViewerLocaleProvider,
  useOfficeFileViewerMessages,
  type OfficeFileViewerLocale,
} from './locale';
import type { OfficeFileViewerUri } from './services/input/normalizeOfficeFileUri';
import type {
  OfficeParseOptions,
  OfficePreviewReadyInfo,
  ParseProgress,
} from './services/parsing';
import {
  isPresentationPreviewKind,
  isSpreadsheetPreviewKind,
  type ParsedOfficeFile,
} from './services/preview';
import { OfficeResourceStoreProvider } from './services/resource-store';
import { OfficeParseStatus } from './shell/ParseStatus';
import {
  OfficePreviewStage,
  type OfficePreviewStageState,
} from './shell/PreviewStage';
import {
  OfficeToolbar,
  type FullscreenControls,
  type OfficeToolbarFormatControls,
  type ZoomControls,
} from './shell/Toolbar';
import { OFFICE_DEFAULT_ZOOM } from './shell/constants';
import { useOfficeViewerController } from './shell/controller/useOfficeViewerController';

const { Content } = Layout;

export type { OfficeFileViewerUri } from './services/input/normalizeOfficeFileUri';

/** Office文件预览器组件属性。 */
export type OfficeFileViewerProps = {
  /** 预览器使用的界面语言，默认使用简体中文。 */
  locale?: OfficeFileViewerLocale;
  /** 待预览文件的来源，可为本地文件、URL 或异步加载函数。 */
  uri?: OfficeFileViewerUri;
  /** 无法从文件来源推断名称时使用的默认文件名。 */
  defaultFileName?: string;
  /** 组件首次渲染时采用的缩放比例。 */
  defaultZoom?: number;
  /** 非受控模式下演讲者备注是否默认展开。 */
  defaultShowSpeakerNotes?: boolean;
  /** 受控模式下演讲者备注是否展开。 */
  showSpeakerNotes?: boolean;
  /** 演讲者备注展开状态变化时触发。 */
  onSpeakerNotesVisibilityChange?: (visible: boolean) => void;
  /** 附加到预览器根元素的自定义类名。 */
  className?: string;
  /** 预览区域高度，支持任意 CSS 高度值；未提供时使用父容器高度。 */
  height?: CSSProperties['height'];
  /** 传递给预览器根元素的内联样式。 */
  style?: CSSProperties;
  /** 完整文件解析成功后触发一次；渐进解析结果不会触发该回调。 */
  onFileParsed?: (parsed: ParsedOfficeFile, file: File) => void;
  /** 首屏预览就绪后触发一次，完整模型和按需数据源都会触发。 */
  onPreviewReady?: (info: OfficePreviewReadyInfo, file: File) => void;
  /** 文件加载或解析失败时触发的回调。 */
  onError?: (error: Error, file?: File) => void;
  /** 传递给底层解析会话的运行配置。 */
  parseOptions?: OfficeParseOptions;
  /** 解析阶段或完成度变化时触发的进度回调。 */
  onParseProgress?: (progress: ParseProgress) => void;
};

/** 组合控制器状态与现有工具栏、预览舞台，保持公共渲染结构稳定。 */
function OfficeFileViewerContent({
  uri,
  defaultFileName,
  defaultZoom = OFFICE_DEFAULT_ZOOM,
  defaultShowSpeakerNotes = false,
  showSpeakerNotes,
  onSpeakerNotesVisibilityChange,
  className,
  height,
  style,
  onFileParsed,
  onPreviewReady,
  onError,
  parseOptions,
  onParseProgress,
}: Omit<OfficeFileViewerProps, 'locale'>) {
  const messages = useOfficeFileViewerMessages();
  const { state, actions, meta, viewerRef, resourceStore } =
    useOfficeViewerController({
      uri,
      defaultZoom,
      defaultShowSpeakerNotes,
      showSpeakerNotes,
      onSpeakerNotesVisibilityChange,
      onFileParsed,
      onPreviewReady,
      onError,
      parseOptions,
      onParseProgress,
      messages,
    });
  const { document: documentState, view } = state;
  const preview = meta.preview;
  const loading =
    documentState.phase === 'resolving' || documentState.phase === 'parsing';
  const parseProgress =
    documentState.phase === 'parsing' ? documentState.progress : undefined;
  const error =
    documentState.phase === 'failed' ? documentState.message : undefined;
  const partialWarning = documentState.phase === 'degraded';
  const loadedFileName =
    documentState.phase === 'parsing' ||
    documentState.phase === 'ready' ||
    documentState.phase === 'degraded' ||
    documentState.phase === 'failed'
      ? documentState.fileName
      : undefined;
  const displayedFileName =
    loadedFileName ?? defaultFileName ?? messages.file.unloaded;
  const hasWordOutline = meta.format.kind === 'word' && meta.format.hasOutline;
  const formatControls = useMemo<OfficeToolbarFormatControls>(() => {
    if (meta.format.kind === 'presentation') {
      return {
        kind: 'presentation',
        navigation:
          meta.format.slideCount > 1
            ? {
                canPrevious: meta.format.canPrevious,
                canNext: meta.format.canNext,
                previous: actions.previousSlide,
                next: actions.nextSlide,
              }
            : undefined,
        speakerNotes: {
          visible: meta.speakerNotesVisible,
          disabled: !meta.hasRenderableContent,
          toggle: actions.toggleSpeakerNotes,
        },
      };
    }
    if (meta.format.kind === 'word') {
      return {
        kind: 'word',
        outline: meta.format.hasOutline
          ? {
              visible: view.showWordOutline,
              toggle: actions.toggleWordOutline,
            }
          : undefined,
      };
    }
    if (meta.format.kind === 'spreadsheet') {
      return { kind: 'spreadsheet' };
    }
    return { kind: 'empty' };
  }, [
    actions.nextSlide,
    actions.previousSlide,
    actions.toggleSpeakerNotes,
    actions.toggleWordOutline,
    meta.format,
    meta.hasRenderableContent,
    meta.speakerNotesVisible,
    view.showWordOutline,
  ]);
  const zoomControls = useMemo<ZoomControls>(
    () => ({
      value: view.zoom,
      hasDocument: meta.hasRenderableContent,
      zoomOut: actions.zoomOut,
      zoomIn: actions.zoomIn,
      change: actions.changeZoom,
      reset: actions.resetZoom,
    }),
    [
      actions.changeZoom,
      actions.resetZoom,
      actions.zoomIn,
      actions.zoomOut,
      meta.hasRenderableContent,
      view.zoom,
    ],
  );
  const fullscreenControls = useMemo<FullscreenControls>(
    () => ({
      active: view.isFullscreen,
      disabled: !meta.hasRenderableContent || !meta.fullscreenSupported,
      toggle: actions.toggleFullscreen,
    }),
    [
      actions.toggleFullscreen,
      meta.fullscreenSupported,
      meta.hasRenderableContent,
      view.isFullscreen,
    ],
  );
  let previewStageState: OfficePreviewStageState;
  if (error) {
    previewStageState = { kind: 'error', message: error };
  } else if (loading && !meta.hasRenderableContent) {
    previewStageState = { kind: 'loading', tip: meta.loadingTip };
  } else if (!preview) {
    previewStageState = { kind: 'empty' };
  } else if (isPresentationPreviewKind(preview.previewKind)) {
    previewStageState = {
      kind: 'presentation',
      preview,
      activeIndex: view.activeSlideIndex,
      zoom: view.zoom,
      showSpeakerNotes: meta.speakerNotesVisible,
    };
  } else if (isSpreadsheetPreviewKind(preview.previewKind)) {
    previewStageState = {
      kind: 'spreadsheet',
      preview,
      activeSheetId: view.activeSheetId,
      zoom: view.zoom,
    };
  } else if (preview.previewKind === 'docx') {
    previewStageState = {
      kind: 'docx',
      preview,
      zoom: view.zoom,
      showOutline: view.showWordOutline && hasWordOutline,
    };
  } else {
    previewStageState = {
      kind: 'doc',
      preview,
      zoom: view.zoom,
      showOutline: view.showWordOutline && hasWordOutline,
    };
  }
  // 专用 height 配置优先于 style.height，避免两个入口同时传值时结果不确定。
  const viewerStyle = height === undefined ? style : { ...style, height };

  return (
    <div
      ref={viewerRef}
      className={['office-file-viewer', className].filter(Boolean).join(' ')}
      style={viewerStyle}
    >
      <Layout className="office-file-viewer__layout">
        <OfficeToolbar
          fileName={displayedFileName}
          previewKind={meta.previewKind}
          formatControls={formatControls}
          zoomControls={zoomControls}
          fullscreenControls={fullscreenControls}
          onSelectFile={actions.selectFile}
        />
        <Content className="office-file-viewer__content">
          <OfficeResourceStoreProvider store={resourceStore}>
            <OfficePreviewStage
              state={previewStageState}
              onCloseWordOutline={actions.closeWordOutline}
              onSelectSlide={actions.selectSlide}
              onSelectSheet={actions.selectSheet}
            />
          </OfficeResourceStoreProvider>
          <OfficeParseStatus
            progress={
              loading && meta.hasRenderableContent ? parseProgress : undefined
            }
            warning={partialWarning}
          />
        </Content>
      </Layout>
    </div>
  );
}

/** 渲染 Office 文件预览器，并为当前实例提供独立的界面语言。 */
function OfficeFileViewerComponent({
  locale = 'zh-CN',
  ...props
}: OfficeFileViewerProps): ReactElement {
  return (
    <OfficeFileViewerLocaleProvider locale={locale}>
      <OfficeFileViewerContent {...props} />
    </OfficeFileViewerLocaleProvider>
  );
}

export const OfficeFileViewer = memo(OfficeFileViewerComponent);
