// OfficeFileViewer 是组件库对外主入口，负责组合本地化、控制器、工具栏和格式预览舞台。
import type { CSSProperties, ReactElement } from 'react';
import React, { memo, useMemo } from 'react';
import { OfficeSearchRuntimeBoundary } from './formats/search/OfficeSearchContext';
import { OfficeSearchSidebar } from './formats/search/OfficeSearchSidebar';
import { useOfficeSearchController } from './formats/search/useOfficeSearchController';
import './index.less';
import {
  OfficeFileViewerLocaleProvider,
  useOfficeFileViewerMessages,
  type OfficeFileViewerLocale,
} from './locale';
import type { OfficeFileViewerFontOptions } from './services/fonts/types';
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
import type { OfficeFileViewerWarning } from './services/previewWarnings';
import { OfficeResourceStoreProvider } from './services/resource-store';
import type { OfficeFileViewerSearchOptions } from './services/search/types';
import { OfficeFontProvider } from './shared/fonts/OfficeFontProvider';
import {
  OfficeHyperlinkProvider,
  type OfficeHyperlinkActivateEvent,
} from './shared/hyperlink';
import {
  OfficeImagePreviewProvider,
  type OfficeFileViewerImagePreviewConfig,
} from './shared/image-preview';
import { OfficeParseStatus } from './shell/ParseStatus';
import {
  OfficePreviewStage,
  type OfficePreviewStageState,
} from './shell/PreviewStage';
import {
  OfficeToolbar,
  type FullscreenControls,
  type OfficeToolbarFormatControls,
  type OfficeToolbarSearchControls,
  type ZoomControls,
} from './shell/Toolbar';
import { OFFICE_DEFAULT_ZOOM } from './shell/constants';
import { useOfficeViewerController } from './shell/controller/useOfficeViewerController';
import type {
  OfficeFileViewerViewState,
  OfficeFileViewerViewStateChange,
} from './shell/viewState';

export type { OfficeFileViewerFontOptions } from './services/fonts/types';
export type {
  OfficeFileViewerUri,
  OfficeFileViewerUriLoader,
} from './services/input/normalizeOfficeFileUri';
export type { OfficeFileViewerSearchOptions } from './services/search/types';

/** 未声明搜索选项时复用稳定空对象，避免无关渲染重置查询。 */
const DEFAULT_SEARCH_OPTIONS: OfficeFileViewerSearchOptions = {};

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
  /** 非受控模式下各视图字段的统一初始值。 */
  defaultViewState?: Partial<OfficeFileViewerViewState>;
  /** 由宿主按字段控制的视图状态。 */
  viewState?: Partial<OfficeFileViewerViewState>;
  /** 用户请求改变视图状态时触发。 */
  onViewStateChange?: (
    state: OfficeFileViewerViewState,
    change: OfficeFileViewerViewStateChange,
  ) => void;
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
  /** 解析降级、格式兼容或运行时诊断警告产生时触发。 */
  onWarning?: (warning: OfficeFileViewerWarning, file: File) => void;
  /** 传递给底层解析会话的运行配置。 */
  parseOptions?: OfficeParseOptions;
  /** 控制内容图片的双击预览、下载和自定义右键菜单，默认全部开启。 */
  imagePreview?: OfficeFileViewerImagePreviewConfig;
  /** 是否启用源文档声明的超链接，默认开启。 */
  hyperlink?: boolean;
  /** 控制文档全文查找；默认启用，传 false 时移除入口和运行时。 */
  search?: false | OfficeFileViewerSearchOptions;
  /** 配置源字体别名、全局回退字体和缺失字体诊断。 */
  fontOptions?: OfficeFileViewerFontOptions;
  /** 链接被有效激活时触发，可阻止组件执行默认导航。 */
  onHyperlinkActivate?: (event: OfficeHyperlinkActivateEvent) => void;
  /** 解析阶段或完成度变化时触发的进度回调。 */
  onParseProgress?: (progress: ParseProgress) => void;
};

/** 组合控制器状态与现有工具栏、预览舞台，保持公共渲染结构稳定。 */
function OfficeFileViewerContent({
  uri,
  defaultFileName,
  defaultZoom = OFFICE_DEFAULT_ZOOM,
  defaultViewState,
  viewState,
  onViewStateChange,
  defaultShowSpeakerNotes = false,
  showSpeakerNotes,
  onSpeakerNotesVisibilityChange,
  className,
  height,
  style,
  onFileParsed,
  onPreviewReady,
  onError,
  onWarning,
  parseOptions,
  imagePreview,
  hyperlink = true,
  search,
  fontOptions,
  onHyperlinkActivate,
  onParseProgress,
}: Omit<OfficeFileViewerProps, 'locale'>) {
  const messages = useOfficeFileViewerMessages();
  const searchEnabled = search !== false;
  const searchOptions = searchEnabled
    ? search ?? DEFAULT_SEARCH_OPTIONS
    : DEFAULT_SEARCH_OPTIONS;
  const { state, actions, meta, viewerRef, resourceStore } =
    useOfficeViewerController({
      uri,
      defaultZoom,
      searchEnabled,
      defaultSearchVisible: Boolean(searchOptions.defaultVisible),
      defaultViewState,
      viewState,
      onViewStateChange,
      defaultShowSpeakerNotes,
      showSpeakerNotes,
      onSpeakerNotesVisibilityChange,
      onFileParsed,
      onPreviewReady,
      onError,
      onWarning,
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
      return {
        kind: 'spreadsheet',
        viewMode: {
          value: view.spreadsheetViewMode,
          disabled: !meta.hasRenderableContent,
          change: actions.changeSpreadsheetViewMode,
        },
      };
    }
    return { kind: 'empty' };
  }, [
    actions.nextSlide,
    actions.previousSlide,
    actions.changeSpreadsheetViewMode,
    actions.toggleSpeakerNotes,
    actions.toggleWordOutline,
    meta.format,
    meta.hasRenderableContent,
    meta.speakerNotesVisible,
    view.showWordOutline,
    view.spreadsheetViewMode,
  ]);
  const zoomControls = useMemo<ZoomControls>(
    () => ({
      value: view.zoom,
      hasDocument: meta.hasRenderableContent,
      zoomOut: actions.zoomOut,
      zoomIn: actions.zoomIn,
      change: actions.changeZoom,
    }),
    [
      actions.changeZoom,
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
  const searchController = useOfficeSearchController({
    enabled: searchEnabled,
    visible: view.showSearch,
    sessionKey: preview?.sessionId,
    options: searchOptions,
    hasDocument: meta.hasRenderableContent,
    onOpen: actions.openSearch,
    onClose: actions.closeSearch,
  });
  const searchControls = useMemo<OfficeToolbarSearchControls>(
    () =>
      searchEnabled
        ? {
            kind: 'enabled',
            visible: view.showSearch,
            disabled: !meta.hasRenderableContent,
            toggle: actions.toggleSearch,
          }
        : { kind: 'disabled' },
    [
      actions.toggleSearch,
      meta.hasRenderableContent,
      searchEnabled,
      view.showSearch,
    ],
  );
  let previewStageState: OfficePreviewStageState;
  if (error) {
    previewStageState = {
      kind: 'error',
      message: error,
      retry: meta.canRetry ? actions.retry : undefined,
    };
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
      viewMode: view.spreadsheetViewMode,
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
      tabIndex={-1}
      onKeyDown={
        searchEnabled ? searchController.handleViewerKeyDown : undefined
      }
    >
      <OfficeResourceStoreProvider store={resourceStore}>
        <OfficeHyperlinkProvider
          containerRef={viewerRef}
          enabled={hyperlink}
          file={meta.currentFile}
          previewKind={meta.previewKind}
          sourceUrl={meta.sourceUrl}
          sessionKey={preview?.sessionId}
          onActivate={onHyperlinkActivate}
          onWarning={onWarning}
        >
          <OfficeImagePreviewProvider
            config={imagePreview}
            sessionKey={preview?.sessionId}
            containerRef={viewerRef}
          >
            <OfficeFontProvider
              options={fontOptions}
              containerRef={viewerRef}
              sessionKey={preview?.sessionId}
              ready={meta.hasRenderableContent}
              file={meta.currentFile}
              previewKind={meta.previewKind}
              onWarning={onWarning}
            >
              <OfficeSearchRuntimeBoundary
                enabled={searchEnabled}
                controller={searchController}
              >
                <div className="office-file-viewer__layout">
                  <OfficeToolbar
                    fileName={displayedFileName}
                    previewKind={meta.previewKind}
                    formatControls={formatControls}
                    zoomControls={zoomControls}
                    fullscreenControls={fullscreenControls}
                    searchControls={searchControls}
                    onSelectFile={actions.selectFile}
                  />
                  <div className="office-file-viewer__content">
                    <div className="office-file-viewer__workspace">
                      {searchEnabled ? (
                        <OfficeSearchSidebar
                          visible={view.showSearch}
                          sessionKey={preview?.sessionId}
                          controller={searchController}
                          onClose={actions.closeSearch}
                          onShowOutline={
                            hasWordOutline
                              ? actions.toggleWordOutline
                              : undefined
                          }
                        />
                      ) : null}
                      <div className="office-file-viewer__stage">
                        <OfficePreviewStage
                          state={previewStageState}
                          onCloseWordOutline={actions.closeWordOutline}
                          onOpenSearch={
                            searchEnabled ? actions.openSearch : undefined
                          }
                          onSelectSlide={actions.selectSlide}
                          onSelectSheet={actions.selectSheet}
                        />
                      </div>
                    </div>
                    <OfficeParseStatus
                      progress={
                        loading && meta.hasRenderableContent
                          ? parseProgress
                          : undefined
                      }
                      warning={partialWarning}
                    />
                  </div>
                </div>
              </OfficeSearchRuntimeBoundary>
            </OfficeFontProvider>
          </OfficeImagePreviewProvider>
        </OfficeHyperlinkProvider>
      </OfficeResourceStoreProvider>
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
