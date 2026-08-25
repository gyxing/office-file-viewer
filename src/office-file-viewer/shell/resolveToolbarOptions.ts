import type {
  OfficeFileViewerToolbarOptions,
  ResolvedOfficeFileViewerToolbarOptions,
} from './Toolbar';

/** 工具栏未声明时显示全部内置操作区域。 */
const DEFAULT_TOOLBAR_OPTIONS: ResolvedOfficeFileViewerToolbarOptions = {
  fileName: true,
  openFile: true,
  wordOutline: true,
  wordRevisionMode: true,
  spreadsheetViewMode: true,
  presentationNavigation: true,
  speakerNotes: true,
  search: true,
  review: true,
  zoom: true,
  fullscreen: true,
};

/** 合并工具栏局部配置，并保持主预览器默认界面向后兼容。 */
export function resolveToolbarOptions(
  options: OfficeFileViewerToolbarOptions | undefined,
): ResolvedOfficeFileViewerToolbarOptions {
  return { ...DEFAULT_TOOLBAR_OPTIONS, ...options };
}
