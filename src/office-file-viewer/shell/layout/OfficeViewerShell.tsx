import type { ComponentType, ReactElement, ReactNode } from 'react';
import React, { createContext, useContext } from 'react';
import type { OfficeCapabilities } from '../../core/types';
import type { OfficeViewerThemeOptions } from '../../shared/theme';
import {
  OfficeWatermarkSurface,
  type OfficeViewerWatermark,
} from '../../shared/watermark';
import { OfficeViewerFrame } from '../frame';
import './index.less';
import {
  OfficeViewerLayoutProvider,
  useOfficeViewerLayout as useLayoutContext,
} from './OfficeViewerLayoutContext';
import type {
  OfficeViewerLayoutContentScaling,
  OfficeViewerShellContextValue,
} from './types';
import { useOfficeViewerLayoutController } from './useOfficeViewerLayoutController';

type OfficeViewerShellProviderBase = {
  theme?: OfficeViewerThemeOptions;
  watermark?: OfficeViewerWatermark;
  capabilities?: OfficeCapabilities;
  children: ReactNode;
};

export type OfficeViewerShellProviderProps = OfficeViewerShellProviderBase &
  (
    | {
        value: OfficeViewerShellContextValue;
        defaultZoom?: never;
        zoom?: never;
        onZoomChange?: never;
        contentScaling?: never;
      }
    | {
        value?: undefined;
        defaultZoom?: number;
        zoom?: number;
        onZoomChange?: (zoom: number) => void;
        contentScaling?: OfficeViewerLayoutContentScaling;
        capabilities?: OfficeCapabilities;
      }
  );

export type OfficeViewerShellComponent = {
  Provider: ComponentType<OfficeViewerShellProviderProps>;
  Root: ComponentType<{ className?: string; children: ReactNode }>;
  Toolbar: ComponentType<{ children?: ReactNode }>;
  Viewport: ComponentType<{ children: ReactNode }>;
  Sidebar: ComponentType<{ children?: ReactNode }>;
  StatusBar: ComponentType<{ children?: ReactNode }>;
  FileInfo: ComponentType<{ children?: ReactNode }>;
  Zoom: ComponentType;
  Fullscreen: ComponentType;
};

type ShellVisualOptions = Pick<
  OfficeViewerShellProviderProps,
  'theme' | 'watermark'
>;
const ShellVisualContext = createContext<ShellVisualOptions>({});

function ShellProvider({
  value,
  defaultZoom,
  zoom,
  onZoomChange,
  contentScaling = 'managed',
  capabilities,
  theme,
  watermark,
  children,
}: OfficeViewerShellProviderProps): ReactElement {
  const controller = useOfficeViewerLayoutController({
    defaultZoom,
    zoom,
    onZoomChange,
    contentScaling,
    capabilities,
  });
  const contextValue = value ?? controller;
  return (
    <OfficeViewerLayoutProvider value={contextValue}>
      <ShellVisualContext.Provider value={{ theme, watermark }}>
        {children}
      </ShellVisualContext.Provider>
    </OfficeViewerLayoutProvider>
  );
}

function ShellRoot({
  className,
  children,
}: {
  className?: string;
  children: ReactNode;
}): ReactElement {
  const context = useLayoutContext();
  const visual = useContext(ShellVisualContext);
  return (
    <OfficeViewerFrame
      viewerRef={context.meta.viewerRef}
      className={[
        'office-viewer-layout',
        'office-viewer-layout--shell',
        className,
      ]
        .filter(Boolean)
        .join(' ')}
      theme={visual.theme}
      watermark={visual.watermark}
    >
      <div className="office-viewer-layout__shell-root">{children}</div>
    </OfficeViewerFrame>
  );
}

const ShellToolbar = ({ children }: { children?: ReactNode }): ReactElement => (
  <div className="office-viewer-layout__shell-toolbar" role="toolbar">
    {children}
  </div>
);
const ShellViewport = ({ children }: { children: ReactNode }): ReactElement => (
  <main className="office-viewer-layout__shell-viewport" tabIndex={0}>
    <OfficeWatermarkSurface className="office-viewer-layout__shell-watermark-surface">
      {children}
    </OfficeWatermarkSurface>
  </main>
);
const ShellSidebar = ({ children }: { children?: ReactNode }): ReactElement => (
  <aside className="office-viewer-layout__shell-sidebar" tabIndex={0}>
    {children}
  </aside>
);
const ShellStatusBar = ({
  children,
}: {
  children?: ReactNode;
}): ReactElement => (
  <footer
    className="office-viewer-layout__shell-status-bar"
    role="status"
    aria-live="polite"
  >
    {children}
  </footer>
);
const ShellFileInfo = ({
  children,
}: {
  children?: ReactNode;
}): ReactElement => (
  <span className="office-viewer-layout__shell-file-info">{children}</span>
);

function ShellZoom(): ReactElement {
  const context = useLayoutContext();
  return (
    <span className="office-viewer-layout__shell-zoom">
      <button type="button" onClick={context.actions.zoomOut} aria-label="缩小">
        −
      </button>
      <span>{context.state.zoom}%</span>
      <button type="button" onClick={context.actions.zoomIn} aria-label="放大">
        +
      </button>
    </span>
  );
}

function ShellFullscreen(): ReactElement {
  const context = useLayoutContext();
  return (
    <button
      type="button"
      onClick={context.actions.toggleFullscreen}
      aria-label="切换全屏"
    >
      {context.state.isFullscreen ? '退出全屏' : '全屏'}
    </button>
  );
}

export const OfficeViewerShell = {
  Provider: ShellProvider,
  Root: ShellRoot,
  Toolbar: ShellToolbar,
  Viewport: ShellViewport,
  Sidebar: ShellSidebar,
  StatusBar: ShellStatusBar,
  FileInfo: ShellFileInfo,
  Zoom: ShellZoom,
  Fullscreen: ShellFullscreen,
} satisfies OfficeViewerShellComponent;
