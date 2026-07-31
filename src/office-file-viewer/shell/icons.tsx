import type { ReactNode, SVGProps } from 'react';
import React from 'react';

/** Office图标组件属性。 */
type OfficeIconProps = Omit<SVGProps<SVGSVGElement>, 'children'>;

/** Office 文件图标基础组件属性。 */
type OfficeIconBaseProps = OfficeIconProps & {
  /** 当前组件包含的子节点。 */
  children: ReactNode;
};

/** Office 文件类型图标共用的文件轮廓路径。 */
const FILE_FRAME = <path d="M6 2.75h7l5 5v13.5H6zM13 2.75v5h5" />;

// OfficeIconBase 统一工具栏图标的尺寸、描边和无障碍属性。
/** 渲染 Office 文件类型图标共用的基础轮廓。 */
function OfficeIconBase({ children, ...props }: OfficeIconBaseProps) {
  return (
    <svg
      {...props}
      aria-hidden="true"
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

/** 渲染尚未识别格式时使用的通用文件图标。 */
export function FileIcon(props: OfficeIconProps) {
  return <OfficeIconBase {...props}>{FILE_FRAME}</OfficeIconBase>;
}

// FileExcelIcon 表示电子表格文件。
/** 渲染 Excel 文件图标。 */
export function FileExcelIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      {FILE_FRAME}
      <path d="M8.5 11h7v6h-7zM12 11v6M8.5 14h7" />
    </OfficeIconBase>
  );
}

// FilePptIcon 表示演示文稿文件。
/** 渲染 PowerPoint 文件图标。 */
export function FilePptIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      {FILE_FRAME}
      <path d="M9 16.5v-5h2.5a1.75 1.75 0 0 1 0 3.5H9" />
    </OfficeIconBase>
  );
}

// FileWordIcon 表示文字文档文件。
/** 渲染 Word 文件图标。 */
export function FileWordIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      {FILE_FRAME}
      <path d="m8.5 11 1.25 6 2.25-4 2.25 4 1.25-6" />
    </OfficeIconBase>
  );
}

// ChevronLeftIcon 表示上一页操作。
/** 渲染向左翻页箭头图标。 */
export function ChevronLeftIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <path d="m15 18-6-6 6-6" />
    </OfficeIconBase>
  );
}

// ChevronRightIcon 表示下一页操作。
/** 渲染向右翻页箭头图标。 */
export function ChevronRightIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <path d="m9 6 6 6-6 6" />
    </OfficeIconBase>
  );
}

// ZoomOutIcon 表示缩小操作。
/** 渲染缩小预览图标。 */
export function ZoomOutIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <circle cx="10.5" cy="10.5" r="6.5" />
      <path d="M7.5 10.5h6M15.5 15.5 21 21" />
    </OfficeIconBase>
  );
}

// ZoomInIcon 表示放大操作。
/** 渲染放大预览图标。 */
export function ZoomInIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <circle cx="10.5" cy="10.5" r="6.5" />
      <path d="M7.5 10.5h6M10.5 7.5v6M15.5 15.5 21 21" />
    </OfficeIconBase>
  );
}

// NotesIcon 表示演讲者备注面板。
/** 渲染演讲者备注图标。 */
export function NotesIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <path d="M5 3.5h14v17H5zM8 8h8M8 12h8M8 16h5" />
    </OfficeIconBase>
  );
}

// FullscreenIcon 表示进入全屏操作。
/** 渲染全屏切换图标。 */
export function FullscreenIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <path d="M8 3H3v5M16 3h5v5M21 16v5h-5M8 21H3v-5" />
    </OfficeIconBase>
  );
}

// OutlineIcon 表示 Word 文档的层级大纲。
/** 渲染文档大纲图标。 */
export function OutlineIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <circle cx="5" cy="7" r="1" />
      <circle cx="5" cy="12" r="1" />
      <circle cx="8" cy="17" r="1" />
      <path d="M9 7h10M9 12h10M12 17h7" />
    </OfficeIconBase>
  );
}

// PanelLeftCloseIcon 表示收起左侧面板。
/** 渲染收起左侧面板图标。 */
export function PanelLeftCloseIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <rect x="3" y="4" width="18" height="16" rx="2" />
      <path d="M8 4v16m7-5-3-3 3-3" />
    </OfficeIconBase>
  );
}

// PanelLeftOpenIcon 表示展开左侧面板。
/** 渲染展开左侧面板图标。 */
export function PanelLeftOpenIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <rect x="3" y="4" width="18" height="16" rx="2" />
      <path d="M8 4v16m4-5 3-3-3-3" />
    </OfficeIconBase>
  );
}
