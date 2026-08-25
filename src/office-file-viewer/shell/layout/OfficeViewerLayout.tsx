import type { CSSProperties, ReactElement } from 'react';
import React, { memo, useMemo } from 'react';
import { OfficeFileViewerLocaleProvider } from '../../locale';
import { OfficeWatermarkSurface } from '../../shared/watermark';
import { OfficeViewerFrame } from '../frame';
import { resolveToolbarOptions } from '../resolveToolbarOptions';
import {
  OfficeToolbar,
  type FullscreenControls,
  type ZoomControls,
} from '../Toolbar';
import './index.less';
import { OfficeViewerLayoutProvider } from './OfficeViewerLayoutContext';
import type { OfficeViewerLayoutProps } from './types';
import { useOfficeViewerLayoutController } from './useOfficeViewerLayoutController';

/** 托管缩放写入内容容器的动态 CSS 变量。 */
interface OfficeViewerLayoutContentStyle extends CSSProperties {
  '--office-viewer-layout-scale': number;
}

/** 通用外壳仅提供百分比缩放，不暴露格式相关适配模式。 */
function keepPercentageZoomMode(): void {}

/** 组合共享工具栏、内容视口和宿主可消费的控制上下文。 */
function OfficeViewerLayoutContent({
  fileName,
  className,
  height,
  style,
  toolbar,
  toolbarExtra,
  fileAccept,
  onFileSelect,
  theme,
  watermark,
  defaultZoom,
  zoom,
  onZoomChange,
  onFullscreenChange,
  onFullscreenError,
  contentScaling = 'managed',
  children,
}: Omit<OfficeViewerLayoutProps, 'locale'>): ReactElement {
  const controller = useOfficeViewerLayoutController({
    defaultZoom,
    zoom,
    onZoomChange,
    onFullscreenChange,
    onFullscreenError,
    contentScaling,
  });
  const { state, actions, meta } = controller;
  const configuredToolbar =
    toolbar === false ? undefined : resolveToolbarOptions(toolbar);
  const displayOptions = configuredToolbar
    ? {
        ...configuredToolbar,
        fileName: configuredToolbar.fileName && Boolean(fileName),
        openFile: configuredToolbar.openFile && Boolean(onFileSelect),
      }
    : undefined;
  const zoomControls = useMemo<ZoomControls>(
    () => ({
      value: state.zoom,
      mode: 'percentage',
      fitModes: [],
      hasDocument: true,
      zoomOut: actions.zoomOut,
      zoomIn: actions.zoomIn,
      change: actions.changeZoom,
      changeMode: keepPercentageZoomMode,
    }),
    [actions, state.zoom],
  );
  const fullscreenControls = useMemo<FullscreenControls>(
    () => ({
      active: state.isFullscreen,
      disabled: !meta.fullscreenSupported,
      toggle: actions.toggleFullscreen,
    }),
    [actions.toggleFullscreen, meta.fullscreenSupported, state.isFullscreen],
  );
  const contentStyle: OfficeViewerLayoutContentStyle | undefined =
    contentScaling === 'managed'
      ? { '--office-viewer-layout-scale': state.zoom / 100 }
      : undefined;

  return (
    <OfficeViewerFrame
      viewerRef={meta.viewerRef}
      className={['office-viewer-layout', className].filter(Boolean).join(' ')}
      height={height}
      style={style}
      theme={theme}
      watermark={watermark}
    >
      <OfficeViewerLayoutProvider value={controller}>
        <div className="office-file-viewer__layout">
          {displayOptions ? (
            <OfficeToolbar
              fileName={fileName ?? ''}
              formatControls={{ kind: 'empty' }}
              zoomControls={zoomControls}
              fullscreenControls={fullscreenControls}
              searchControls={{ kind: 'disabled' }}
              reviewControls={{ kind: 'disabled' }}
              displayOptions={displayOptions}
              extra={toolbarExtra}
              fileAccept={fileAccept}
              onSelectFile={onFileSelect}
            />
          ) : null}
          <div className="office-file-viewer__content">
            <OfficeWatermarkSurface className="office-viewer-layout__watermark-surface">
              <div className="office-viewer-layout__viewport">
                <div
                  className="office-viewer-layout__content"
                  data-content-scaling={contentScaling}
                  style={contentStyle}
                >
                  {children}
                </div>
              </div>
            </OfficeWatermarkSurface>
          </div>
        </div>
      </OfficeViewerLayoutProvider>
    </OfficeViewerFrame>
  );
}

/** 渲染可复用预览外壳，并提供独立语言环境。 */
function OfficeViewerLayoutComponent({
  locale = 'zh-CN',
  ...props
}: OfficeViewerLayoutProps): ReactElement {
  return (
    <OfficeFileViewerLocaleProvider locale={locale}>
      <OfficeViewerLayoutContent {...props} />
    </OfficeFileViewerLocaleProvider>
  );
}

export const OfficeViewerLayout = memo(OfficeViewerLayoutComponent);
