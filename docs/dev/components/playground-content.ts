import type { WebsiteLocale } from './home-content';

export type PlaygroundTarget = 'viewer' | 'layout';
export type PlaygroundToolbarMode = 'default' | 'custom' | 'hidden';
export type PlaygroundCodeTab = 'tsx' | 'props';

export type PlaygroundContent = {
  eyebrow: string;
  title: string;
  description: string;
  privacy: string;
  controlsTitle: string;
  reset: string;
  component: string;
  targetViewer: string;
  targetLayout: string;
  appearance: string;
  themeMode: string;
  themeLight: string;
  themeDark: string;
  themeSystem: string;
  primaryColor: string;
  workspaceColor: string;
  watermark: string;
  watermarkEnabled: string;
  watermarkContent: string;
  watermarkColor: string;
  watermarkOpacity: string;
  watermarkRotate: string;
  toolbar: string;
  toolbarMode: string;
  toolbarDefault: string;
  toolbarCustom: string;
  toolbarHidden: string;
  toolbarFileName: string;
  toolbarOpenFile: string;
  toolbarWordOutline: string;
  toolbarWordRevisionMode: string;
  toolbarSpreadsheetViewMode: string;
  toolbarPresentationNavigation: string;
  toolbarSpeakerNotes: string;
  toolbarSearch: string;
  toolbarReview: string;
  toolbarZoom: string;
  toolbarFullscreen: string;
  preview: string;
  zoom: string;
  previewHeight: string;
  search: string;
  review: string;
  imagePreview: string;
  layoutBehavior: string;
  controlledZoom: string;
  contentScaling: string;
  managedScaling: string;
  manualScaling: string;
  previewTitle: string;
  viewerHint: string;
  layoutHint: string;
  selectedFile: string;
  noFile: string;
  exampleCode: string;
  codeTab: string;
  propsTab: string;
  copy: string;
  copied: string;
  copyFailed: string;
  customContentTitle: string;
  customContentDescription: string;
  runtimeZoom: string;
  runtimeFullscreen: string;
  runtimeScaling: string;
  yes: string;
  no: string;
  layoutCardInteraction: string;
  layoutCardAppearance: string;
  layoutCardExtensibility: string;
};

const PLAYGROUND_CONTENT: Record<WebsiteLocale, PlaygroundContent> = {
  'en-US': {
    eyebrow: 'Interactive playground',
    title: 'Tune the viewer and use the generated code.',
    description:
      'Switch components, themes, watermarks, toolbars, and viewing behavior. Every change is reflected in the preview and example code immediately.',
    privacy:
      'Files selected here stay in this browser tab and are not uploaded.',
    controlsTitle: 'Parameters',
    reset: 'Reset defaults',
    component: 'Component',
    targetViewer: 'OfficeFileViewer',
    targetLayout: 'OfficeViewerLayout',
    appearance: 'Appearance',
    themeMode: 'Theme mode',
    themeLight: 'Light',
    themeDark: 'Dark',
    themeSystem: 'System',
    primaryColor: 'Primary color',
    workspaceColor: 'Workspace',
    watermark: 'Watermark',
    watermarkEnabled: 'Show watermark',
    watermarkContent: 'Text',
    watermarkColor: 'Color',
    watermarkOpacity: 'Opacity',
    watermarkRotate: 'Rotation',
    toolbar: 'Toolbar',
    toolbarMode: 'Toolbar mode',
    toolbarDefault: 'Default',
    toolbarCustom: 'Custom',
    toolbarHidden: 'Hidden',
    toolbarFileName: 'File name',
    toolbarOpenFile: 'Open file',
    toolbarWordOutline: 'Word outline',
    toolbarWordRevisionMode: 'Revision view',
    toolbarSpreadsheetViewMode: 'Spreadsheet view mode',
    toolbarPresentationNavigation: 'Slide navigation',
    toolbarSpeakerNotes: 'Speaker notes',
    toolbarSearch: 'Search action',
    toolbarReview: 'Review action',
    toolbarZoom: 'Zoom',
    toolbarFullscreen: 'Fullscreen',
    preview: 'Viewing',
    zoom: 'Initial zoom',
    previewHeight: 'Preview height',
    search: 'Search',
    review: 'Review',
    imagePreview: 'Image preview',
    layoutBehavior: 'Layout behavior',
    controlledZoom: 'Controlled zoom',
    contentScaling: 'Content scaling',
    managedScaling: 'Managed',
    manualScaling: 'Manual',
    previewTitle: 'Live preview',
    viewerHint: 'Open a local Office file to test the real viewer.',
    layoutHint:
      'This mode wraps host-rendered content and does not parse the selected file.',
    selectedFile: 'Selected file',
    noFile: 'No local file selected',
    exampleCode: 'Example code',
    codeTab: 'React / TSX',
    propsTab: 'Props',
    copy: 'Copy',
    copied: 'Copied',
    copyFailed: 'Copy manually',
    customContentTitle: 'Host content preview',
    customContentDescription:
      'The layout supplies a shared toolbar, zoom, fullscreen, theme, and watermark while your application owns the content.',
    runtimeZoom: 'Zoom',
    runtimeFullscreen: 'Fullscreen',
    runtimeScaling: 'Scaling',
    yes: 'Yes',
    no: 'No',
    layoutCardInteraction: 'Shared interaction',
    layoutCardAppearance: 'Shared appearance',
    layoutCardExtensibility: 'Open composition',
  },
  'zh-CN': {
    eyebrow: '在线体验',
    title: '调整预览参数，并直接使用生成的代码。',
    description:
      '切换组件、主题、水印、工具栏和查看行为，所有修改都会立即同步到预览效果与示例代码。',
    privacy: '这里选择的文件只保留在当前浏览器标签页，不会被上传。',
    controlsTitle: '组件参数',
    reset: '恢复默认值',
    component: '目标组件',
    targetViewer: 'OfficeFileViewer',
    targetLayout: 'OfficeViewerLayout',
    appearance: '外观',
    themeMode: '主题模式',
    themeLight: '浅色',
    themeDark: '深色',
    themeSystem: '跟随系统',
    primaryColor: '主题主色',
    workspaceColor: '工作区背景',
    watermark: '水印',
    watermarkEnabled: '显示水印',
    watermarkContent: '水印文字',
    watermarkColor: '水印颜色',
    watermarkOpacity: '透明度',
    watermarkRotate: '旋转角度',
    toolbar: '工具栏',
    toolbarMode: '工具栏模式',
    toolbarDefault: '默认配置',
    toolbarCustom: '自定义',
    toolbarHidden: '隐藏',
    toolbarFileName: '文件名',
    toolbarOpenFile: '打开文件',
    toolbarWordOutline: 'Word 大纲',
    toolbarWordRevisionMode: '修订视图',
    toolbarSpreadsheetViewMode: '表格显示模式',
    toolbarPresentationNavigation: 'PPT 翻页',
    toolbarSpeakerNotes: '演讲者备注',
    toolbarSearch: '查找入口',
    toolbarReview: '审阅入口',
    toolbarZoom: '缩放',
    toolbarFullscreen: '全屏',
    preview: '查看行为',
    zoom: '初始缩放',
    previewHeight: '预览高度',
    search: '全文查找',
    review: '审阅标记',
    imagePreview: '图片预览',
    layoutBehavior: '外壳行为',
    controlledZoom: '受控缩放',
    contentScaling: '内容缩放',
    managedScaling: '外壳托管',
    manualScaling: '宿主处理',
    previewTitle: '实时预览',
    viewerHint: '打开一个本地 Office 文件，体验真实预览组件。',
    layoutHint: '该模式包装宿主渲染的内容，不会解析所选文件。',
    selectedFile: '当前文件',
    noFile: '尚未选择本地文件',
    exampleCode: '示例代码',
    codeTab: 'React / TSX',
    propsTab: '参数配置',
    copy: '复制代码',
    copied: '已复制',
    copyFailed: '请手动复制',
    customContentTitle: '宿主内容预览',
    customContentDescription:
      '外壳统一提供工具栏、缩放、全屏、主题与水印，实际内容仍由宿主项目自由渲染。',
    runtimeZoom: '当前缩放',
    runtimeFullscreen: '全屏状态',
    runtimeScaling: '缩放职责',
    yes: '是',
    no: '否',
    layoutCardInteraction: '统一交互',
    layoutCardAppearance: '统一外观',
    layoutCardExtensibility: '开放组合',
  },
};

/** 返回与在线体验页语言一致的参数和预览文案。 */
export function getPlaygroundContent(locale: WebsiteLocale): PlaygroundContent {
  return PLAYGROUND_CONTENT[locale];
}
