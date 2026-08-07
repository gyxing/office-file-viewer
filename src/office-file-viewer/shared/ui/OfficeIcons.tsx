import type { ReactNode, SVGProps } from 'react';
import React from 'react';

/** 预览器通用 SVG 图标可接收的属性。 */
type OfficeIconProps = Omit<SVGProps<SVGSVGElement>, 'children'>;

/** 线性 SVG 图标基础组件属性。 */
type OfficeIconBaseProps = OfficeIconProps & {
  /** 当前组件包含的子节点。 */
  children: ReactNode;
};

/** 复制既有打开文件图标的矢量路径，避免为单个图标新增运行时依赖。 */
const FOLDER_OPEN_OUTLINED_PATH =
  'M928 444H820V330.4c0-17.7-14.3-32-32-32H473L355.7 186.2a8.15 8.15 0 00-5.5-2.2H96c-17.7 0-32 14.3-32 32v592c0 17.7 14.3 32 32 32h698c13 0 24.8-7.9 29.7-20l134-332c1.5-3.8 2.3-7.9 2.3-12 0-17.7-14.3-32-32-32zM136 256h188.5l119.6 114.4H748V444H238c-13 0-24.8 7.9-29.7 20L136 643.2V256zm635.3 512H159l103.3-256h612.4L771.3 768z';

/** 复制既有大纲图标的矢量路径，供工具栏和大纲面板复用。 */
const BARS_OUTLINED_PATH =
  'M912 192H328c-4.4 0-8 3.6-8 8v56c0 4.4 3.6 8 8 8h584c4.4 0 8-3.6 8-8v-56c0-4.4-3.6-8-8-8zm0 284H328c-4.4 0-8 3.6-8 8v56c0 4.4 3.6 8 8 8h584c4.4 0 8-3.6 8-8v-56c0-4.4-3.6-8-8-8zm0 284H328c-4.4 0-8 3.6-8 8v56c0 4.4 3.6 8 8 8h584c4.4 0 8-3.6 8-8v-56c0-4.4-3.6-8-8-8zM104 228a56 56 0 10112 0 56 56 0 10-112 0zm0 284a56 56 0 10112 0 56 56 0 10-112 0zm0 284a56 56 0 10112 0 56 56 0 10-112 0z';

/** 统一线性图标的 SVG 画布、描边和无障碍属性。 */
function OfficeIconBase({
  children,
  className,
  ...props
}: OfficeIconBaseProps) {
  const mergedClassName = ['office-file-icon', className]
    .filter(Boolean)
    .join(' ');

  return (
    <svg
      {...props}
      aria-hidden="true"
      className={mergedClassName}
      focusable="false"
      width="1em"
      height="1em"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      {children}
    </svg>
  );
}

/** 渲染打开文件图标。 */
export function FolderOpenIcon({ className, ...props }: OfficeIconProps) {
  const mergedClassName = ['office-file-icon', className]
    .filter(Boolean)
    .join(' ');

  return (
    <svg
      {...props}
      aria-hidden="true"
      className={mergedClassName}
      focusable="false"
      width="1em"
      height="1em"
      viewBox="64 64 896 896"
      fill="currentColor"
    >
      <path d={FOLDER_OPEN_OUTLINED_PATH} />
    </svg>
  );
}

/** 渲染向左翻页箭头图标。 */
export function ChevronLeftIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <path d="m15 18-6-6 6-6" />
    </OfficeIconBase>
  );
}

/** 渲染向右翻页箭头图标。 */
export function ChevronRightIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <path d="m9 6 6 6-6 6" />
    </OfficeIconBase>
  );
}

/** 渲染缩小预览图标。 */
export function ZoomOutIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <circle cx="10.5" cy="10.5" r="6.5" />
      <path d="M7.5 10.5h6M15.5 15.5 21 21" />
    </OfficeIconBase>
  );
}

/** 渲染放大预览图标。 */
export function ZoomInIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <circle cx="10.5" cy="10.5" r="6.5" />
      <path d="M7.5 10.5h6M10.5 7.5v6M15.5 15.5 21 21" />
    </OfficeIconBase>
  );
}

/** 渲染演讲者备注图标。 */
export function NotesIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <path d="M5 3.5h14v17H5zM8 8h8M8 12h8M8 16h5" />
    </OfficeIconBase>
  );
}

/** 渲染全屏切换图标。 */
export function FullscreenIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <path d="M8 3H3v5M16 3h5v5M21 16v5h-5M8 21H3v-5" />
    </OfficeIconBase>
  );
}

/** 渲染文档大纲图标。 */
export function OutlineIcon({ className, ...props }: OfficeIconProps) {
  const mergedClassName = ['office-file-icon', className]
    .filter(Boolean)
    .join(' ');

  return (
    <svg
      {...props}
      aria-hidden="true"
      className={mergedClassName}
      focusable="false"
      width="1em"
      height="1em"
      viewBox="0 0 1024 1024"
      fill="currentColor"
    >
      <path d={BARS_OUTLINED_PATH} />
    </svg>
  );
}

/** 渲染收起左侧面板图标。 */
export function PanelLeftCloseIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <rect x="3" y="4" width="18" height="16" rx="2" />
      <path d="M8 4v16m7-5-3-3 3-3" />
    </OfficeIconBase>
  );
}

/** 渲染展开左侧面板图标。 */
export function PanelLeftOpenIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <rect x="3" y="4" width="18" height="16" rx="2" />
      <path d="M8 4v16m4-5 3-3-3-3" />
    </OfficeIconBase>
  );
}
