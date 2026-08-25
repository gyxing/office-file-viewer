import type {
  OfficeFileViewerToolbarOptions,
  OfficeViewerLayoutContentScaling,
  OfficeViewerThemeOptions,
  OfficeViewerWatermark,
} from 'office-file-viewer';
import type {
  PlaygroundTarget,
  PlaygroundToolbarMode,
} from './playground-content';

/** 在线体验页可即时调整的稳定参数集合。 */
export type PlaygroundConfig = {
  /** 当前演示完整预览器还是可复用外壳。 */
  target: PlaygroundTarget;
  /** 预览外壳采用的主题模式。 */
  themeMode: 'light' | 'dark' | 'system';
  /** 工具栏与选中状态使用的主题主色。 */
  primaryColor: string;
  /** 文档内容区外围的工作区背景色。 */
  workspaceColor: string;
  /** 是否在内容视口上方显示水印。 */
  watermarkEnabled: boolean;
  /** 水印重复显示的文字。 */
  watermarkContent: string;
  /** 水印文字颜色。 */
  watermarkColor: string;
  /** 水印透明度。 */
  watermarkOpacity: number;
  /** 水印顺时针旋转角度。 */
  watermarkRotate: number;
  /** 工具栏使用默认、自定义或隐藏配置。 */
  toolbarMode: PlaygroundToolbarMode;
  /** 自定义工具栏是否显示文件名。 */
  toolbarFileName: boolean;
  /** 自定义工具栏是否显示打开文件入口。 */
  toolbarOpenFile: boolean;
  /** 自定义工具栏是否显示 Word 大纲入口。 */
  toolbarWordOutline: boolean;
  /** 自定义工具栏是否显示 Word 修订投影切换。 */
  toolbarWordRevisionMode: boolean;
  /** 自定义工具栏是否显示电子表格显示模式切换。 */
  toolbarSpreadsheetViewMode: boolean;
  /** 自定义工具栏是否显示演示文稿翻页操作。 */
  toolbarPresentationNavigation: boolean;
  /** 自定义工具栏是否显示演讲者备注入口。 */
  toolbarSpeakerNotes: boolean;
  /** 自定义工具栏是否显示全文查找入口。 */
  toolbarSearch: boolean;
  /** 自定义工具栏是否显示文档审阅入口。 */
  toolbarReview: boolean;
  /** 自定义工具栏是否显示缩放操作。 */
  toolbarZoom: boolean;
  /** 自定义工具栏是否显示全屏操作。 */
  toolbarFullscreen: boolean;
  /** 预览内容的缩放百分比。 */
  zoom: number;
  /** 当前预览区高度，单位为像素。 */
  previewHeight: number;
  /** 完整预览器是否启用全文查找。 */
  searchEnabled: boolean;
  /** 完整预览器是否启用批注与修订审阅。 */
  reviewEnabled: boolean;
  /** 完整预览器是否启用内容图片预览。 */
  imagePreviewEnabled: boolean;
  /** 外壳缩放是否交由宿主状态控制。 */
  layoutControlledZoom: boolean;
  /** 外壳还是宿主内容负责应用缩放。 */
  layoutContentScaling: OfficeViewerLayoutContentScaling;
};

/** 在线体验页首次加载与恢复操作使用的统一默认值。 */
export const DEFAULT_PLAYGROUND_CONFIG: PlaygroundConfig = {
  target: 'viewer',
  themeMode: 'light',
  primaryColor: '#1677ff',
  workspaceColor: '#eef1f6',
  watermarkEnabled: false,
  watermarkContent: 'Office File Viewer',
  watermarkColor: '#64748b',
  watermarkOpacity: 0.14,
  watermarkRotate: -22,
  toolbarMode: 'default',
  toolbarFileName: true,
  toolbarOpenFile: true,
  toolbarWordOutline: true,
  toolbarWordRevisionMode: true,
  toolbarSpreadsheetViewMode: true,
  toolbarPresentationNavigation: true,
  toolbarSpeakerNotes: true,
  toolbarSearch: true,
  toolbarReview: true,
  toolbarZoom: true,
  toolbarFullscreen: true,
  zoom: 100,
  previewHeight: 640,
  searchEnabled: true,
  reviewEnabled: true,
  imagePreviewEnabled: true,
  layoutControlledZoom: false,
  layoutContentScaling: 'managed',
};

/** 将页面输入收敛为预览器和复用外壳共用的主题契约。 */
export function resolvePlaygroundTheme(
  config: PlaygroundConfig,
): OfficeViewerThemeOptions {
  return {
    mode: config.themeMode,
    primaryColor: config.primaryColor,
    tokens: { workspaceColor: config.workspaceColor },
  };
}

/** 仅在启用后生成水印参数，避免关闭状态仍创建无效图案。 */
export function resolvePlaygroundWatermark(
  config: PlaygroundConfig,
): OfficeViewerWatermark {
  if (!config.watermarkEnabled) return false;

  return {
    content: config.watermarkContent,
    color: config.watermarkColor,
    opacity: config.watermarkOpacity,
    rotate: config.watermarkRotate,
  };
}

/** 根据三态工具栏选择返回默认值、自定义区域或完整隐藏。 */
export function resolvePlaygroundToolbar(
  config: PlaygroundConfig,
): false | OfficeFileViewerToolbarOptions | undefined {
  if (config.toolbarMode === 'hidden') return false;
  if (config.toolbarMode === 'default') return undefined;

  const commonOptions: OfficeFileViewerToolbarOptions = {
    fileName: config.toolbarFileName,
    openFile: config.toolbarOpenFile,
    zoom: config.toolbarZoom,
    fullscreen: config.toolbarFullscreen,
  };

  if (config.target === 'layout') return commonOptions;

  return {
    ...commonOptions,
    wordOutline: config.toolbarWordOutline,
    wordRevisionMode: config.toolbarWordRevisionMode,
    spreadsheetViewMode: config.toolbarSpreadsheetViewMode,
    presentationNavigation: config.toolbarPresentationNavigation,
    speakerNotes: config.toolbarSpeakerNotes,
    search: config.toolbarSearch,
    review: config.toolbarReview,
  };
}
