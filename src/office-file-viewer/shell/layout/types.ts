import type { CSSProperties, ReactNode, RefObject } from 'react';
import type { OfficeFileViewerLocale } from '../../locale';
import type { OfficeViewerThemeOptions } from '../../shared/theme';
import type { OfficeViewerWatermark } from '../../shared/watermark';
import type { OfficeFileViewerToolbarOptions } from '../Toolbar';

/** 宿主内容采用的缩放职责。 */
export type OfficeViewerLayoutContentScaling = 'managed' | 'manual';

/** 可复用预览外壳的公共视图状态。 */
export type OfficeViewerLayoutState = Readonly<{
  /** 当前缩放百分比。 */
  zoom: number;
  /** 外壳根节点当前是否处于浏览器全屏。 */
  isFullscreen: boolean;
}>;

/** 宿主内容和自定义操作可调用的外壳动作。 */
export type OfficeViewerLayoutActions = Readonly<{
  /** 设置缩放百分比。 */
  changeZoom(value: number): void;
  /** 按内置步长缩小。 */
  zoomOut(): void;
  /** 按内置步长放大。 */
  zoomIn(): void;
  /** 进入或退出浏览器全屏。 */
  toggleFullscreen(): void;
}>;

/** 外壳运行环境信息。 */
export type OfficeViewerLayoutMeta = Readonly<{
  /** 外壳根节点引用。 */
  viewerRef: RefObject<HTMLDivElement>;
  /** 当前浏览器是否支持标准全屏 API。 */
  fullscreenSupported: boolean;
  /** 当前由外壳还是宿主负责应用缩放。 */
  contentScaling: OfficeViewerLayoutContentScaling;
}>;

/** `useOfficeViewerLayout` 返回的稳定上下文契约。 */
export type OfficeViewerLayoutContextValue = Readonly<{
  /** 当前视图状态。 */
  state: OfficeViewerLayoutState;
  /** 可调用的视图动作。 */
  actions: OfficeViewerLayoutActions;
  /** 浏览器能力与根节点引用。 */
  meta: OfficeViewerLayoutMeta;
}>;

/** 可复用预览外壳组件属性。 */
export type OfficeViewerLayoutProps = {
  /** 外壳使用的界面语言，默认使用简体中文。 */
  locale?: OfficeFileViewerLocale;
  /** 工具栏显示的宿主文件名；为空时隐藏文件信息区。 */
  fileName?: string;
  /** 根元素附加类名。 */
  className?: string;
  /** 外壳高度，支持任意 CSS 高度值。 */
  height?: CSSProperties['height'];
  /** 根元素内联样式，其中 CSS 变量优先于 theme。 */
  style?: CSSProperties;
  /** 控制通用工具栏；传 false 时完全隐藏。 */
  toolbar?: false | OfficeFileViewerToolbarOptions;
  /** 追加到全部内置操作之后的宿主工具栏内容。 */
  toolbarExtra?: ReactNode;
  /** 文件选择器接受的文件类型。 */
  fileAccept?: string;
  /** 用户通过内置入口选择文件时触发；为空时隐藏打开文件入口。 */
  onFileSelect?: (file: File) => void;
  /** 外壳主题配置，不改变 children 内部的文档样式。 */
  theme?: OfficeViewerThemeOptions;
  /** 覆盖宿主内容视口的水印配置。 */
  watermark?: OfficeViewerWatermark;
  /** 非受控缩放的初始值，默认 100。 */
  defaultZoom?: number;
  /** 由宿主控制的缩放百分比。 */
  zoom?: number;
  /** 用户请求改变缩放时触发。 */
  onZoomChange?: (zoom: number) => void;
  /** 浏览器全屏状态改变时触发。 */
  onFullscreenChange?: (fullscreen: boolean) => void;
  /** 请求全屏失败时触发。 */
  onFullscreenError?: (error: Error) => void;
  /** 缩放由外壳自动应用还是交给宿主处理，默认 managed。 */
  contentScaling?: OfficeViewerLayoutContentScaling;
  /** 宿主提供的文档内容。 */
  children: ReactNode;
};
