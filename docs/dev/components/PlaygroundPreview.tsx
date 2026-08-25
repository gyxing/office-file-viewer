import {
  OfficeFileViewer,
  OfficeViewerLayout,
  useOfficeViewerLayout,
  type OfficeFileViewerViewStateChange,
  type OfficeViewerLayoutContentScaling,
} from 'office-file-viewer';
import type { CSSProperties } from 'react';
import React, { useEffect } from 'react';
import type { WebsiteLocale } from './home-content';
import type { PlaygroundConfig } from './playground-config';
import {
  resolvePlaygroundTheme,
  resolvePlaygroundToolbar,
  resolvePlaygroundWatermark,
} from './playground-config';
import type { PlaygroundContent } from './playground-content';

type PlaygroundPreviewProps = {
  /** 当前在线体验参数。 */
  config: PlaygroundConfig;
  /** 当前语言的预览文案。 */
  content: PlaygroundContent;
  /** 组件使用的界面语言。 */
  locale: WebsiteLocale;
  /** 用户在体验页选择的本地文件。 */
  selectedFile?: File;
  /** 用户选择新文件后的回调。 */
  onFileSelect(file: File): void;
  /** 工具栏请求改变缩放后的回调。 */
  onZoomChange(zoom: number): void;
};

type PlaygroundLayoutContentProps = {
  /** 宿主内容采用的缩放职责。 */
  contentScaling: OfficeViewerLayoutContentScaling;
  /** 左侧参数区请求同步的缩放值。 */
  requestedZoom: number;
  /** 当前语言的宿主内容文案。 */
  content: PlaygroundContent;
};

/** 在手动模式中由宿主应用缩放，并展示外壳 Hook 的实时状态。 */
function PlaygroundLayoutContent({
  contentScaling,
  requestedZoom,
  content,
}: PlaygroundLayoutContentProps) {
  const { state, actions, meta } = useOfficeViewerLayout();
  const scale = state.zoom / 100;
  const manualStyle: CSSProperties | undefined =
    contentScaling === 'manual'
      ? {
          transform: `scale(${scale})`,
          transformOrigin: 'top left',
          width: `${100 / scale}%`,
        }
      : undefined;

  useEffect(() => {
    // 非受控示例仍需响应左侧参数输入，因此通过公开动作同步外部请求值。
    if (state.zoom !== requestedZoom) actions.changeZoom(requestedZoom);
  }, [actions, requestedZoom, state.zoom]);

  return (
    <div className="office-viewer-playground-host-stage" style={manualStyle}>
      <article className="office-viewer-playground-host-document">
        <header>
          <div>
            <span>OFFICE VIEWER LAYOUT</span>
            <h3>{content.customContentTitle}</h3>
            <p>{content.customContentDescription}</p>
          </div>
          <strong>Layout API</strong>
        </header>
        <section className="office-viewer-playground-runtime">
          <div>
            <span>{content.runtimeZoom}</span>
            <strong>{state.zoom}%</strong>
          </div>
          <div>
            <span>{content.runtimeFullscreen}</span>
            <strong>{state.isFullscreen ? content.yes : content.no}</strong>
          </div>
          <div>
            <span>{content.runtimeScaling}</span>
            <strong>{meta.contentScaling}</strong>
          </div>
        </section>
        <section className="office-viewer-playground-host-cards">
          <div>
            <span>01</span>
            <strong>{content.layoutCardInteraction}</strong>
            <p>Zoom · Fullscreen · Open file</p>
          </div>
          <div>
            <span>02</span>
            <strong>{content.layoutCardAppearance}</strong>
            <p>Theme · Watermark · Toolbar</p>
          </div>
          <div>
            <span>03</span>
            <strong>{content.layoutCardExtensibility}</strong>
            <p>State · Actions · ReactNode</p>
          </div>
        </section>
        <div className="office-viewer-playground-host-actions">
          <button type="button" onClick={actions.zoomOut}>
            −
          </button>
          <button type="button" onClick={actions.zoomIn}>
            +
          </button>
        </div>
      </article>
    </div>
  );
}

/** 渲染当前目标组件，并将工具栏交互回写到同一组页面参数。 */
export function PlaygroundPreview({
  config,
  content,
  locale,
  selectedFile,
  onFileSelect,
  onZoomChange,
}: PlaygroundPreviewProps) {
  const theme = resolvePlaygroundTheme(config);
  const watermark = resolvePlaygroundWatermark(config);
  const toolbar = resolvePlaygroundToolbar(config);
  const fileName = selectedFile?.name ?? content.noFile;
  const handleViewStateChange = (
    _state: unknown,
    change: OfficeFileViewerViewStateChange,
  ) => {
    if (change.key === 'zoom') onZoomChange(change.value);
  };

  if (config.target === 'viewer') {
    return (
      <OfficeFileViewer
        locale={locale}
        height={config.previewHeight}
        toolbar={toolbar}
        theme={theme}
        watermark={watermark}
        viewState={{ zoom: config.zoom }}
        onViewStateChange={handleViewStateChange}
        onFileSelect={onFileSelect}
        search={config.searchEnabled ? undefined : false}
        review={config.reviewEnabled ? undefined : false}
        imagePreview={config.imagePreviewEnabled ? undefined : false}
      />
    );
  }

  return (
    <OfficeViewerLayout
      locale={locale}
      fileName={fileName}
      height={config.previewHeight}
      toolbar={toolbar}
      theme={theme}
      watermark={watermark}
      defaultZoom={config.zoom}
      zoom={config.layoutControlledZoom ? config.zoom : undefined}
      onZoomChange={onZoomChange}
      contentScaling={config.layoutContentScaling}
      fileAccept="*/*"
      onFileSelect={onFileSelect}
    >
      <PlaygroundLayoutContent
        contentScaling={config.layoutContentScaling}
        requestedZoom={config.zoom}
        content={content}
      />
    </OfficeViewerLayout>
  );
}
