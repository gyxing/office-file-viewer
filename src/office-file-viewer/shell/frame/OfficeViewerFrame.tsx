import type {
  CSSProperties,
  KeyboardEventHandler,
  ReactElement,
  ReactNode,
  RefObject,
} from 'react';
import React from 'react';
import '../../index.less';
import {
  resolveOfficeViewerThemeMode,
  resolveOfficeViewerThemeStyle,
  type OfficeViewerThemeOptions,
} from '../../shared/theme';
import {
  OfficeWatermarkProvider,
  type OfficeViewerWatermark,
} from '../../shared/watermark';

/** 共享预览根容器属性。 */
type OfficeViewerFrameProps = {
  /** 根容器引用，用于全屏与快捷键范围判断。 */
  viewerRef: RefObject<HTMLDivElement>;
  /** 根容器附加类名。 */
  className?: string;
  /** 根容器高度，优先于 style.height。 */
  height?: CSSProperties['height'];
  /** 根容器自定义样式；其中 CSS 变量优先于 theme。 */
  style?: CSSProperties;
  /** 外壳主题配置。 */
  theme?: OfficeViewerThemeOptions;
  /** 内容视口水印配置。 */
  watermark?: OfficeViewerWatermark;
  /** 根容器键盘事件。 */
  onKeyDown?: KeyboardEventHandler<HTMLDivElement>;
  /** 共享外壳内部结构。 */
  children: ReactNode;
};

/** 统一主题变量、全屏根节点和水印上下文，不干预源文档内容样式。 */
export function OfficeViewerFrame({
  viewerRef,
  className,
  height,
  style,
  theme,
  watermark,
  onKeyDown,
  children,
}: OfficeViewerFrameProps): ReactElement {
  const frameStyle = {
    ...resolveOfficeViewerThemeStyle(theme),
    ...style,
    ...(height === undefined ? undefined : { height }),
  };

  return (
    <div
      ref={viewerRef}
      className={['office-file-viewer', className].filter(Boolean).join(' ')}
      data-office-theme={resolveOfficeViewerThemeMode(theme)}
      style={frameStyle}
      tabIndex={-1}
      onKeyDown={onKeyDown}
    >
      <OfficeWatermarkProvider watermark={watermark}>
        {children}
      </OfficeWatermarkProvider>
    </div>
  );
}
