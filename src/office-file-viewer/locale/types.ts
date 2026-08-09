import type { ParseStage } from '../services/parsing';
import type { PreviewKind } from '../services/preview';

/** OfficeFileViewer 内置支持的界面语言。 */
export type OfficeFileViewerLocale = 'zh-CN' | 'en-US';

/** 定义预览器内部全部用户可见文案，保证两种语言结构一致。 */
export type OfficeFileViewerMessages = {
  /** 文件读取和格式识别相关文案。 */
  file: {
    /** 尚未选择文件时显示的文案。 */
    unloaded: string;
    /** 引导用户选择文件开始预览的文案。 */
    selectToPreview: string;
    /** 文件格式不受支持时显示的文案。 */
    unsupported: string;
    /** 无法识别文件格式时显示的文案。 */
    unrecognized: string;
    /** 文件地址无效时显示的文案。 */
    invalidUri: string;
    /** 文件解析失败时显示的文案。 */
    parseFailed: string;
    /** 文件加载失败时显示的文案。 */
    loadFailed: string;
    /** 文件下载失败时生成提示文案的函数。 */
    downloadFailed: (status: number, statusText: string) => string;
    /** 浏览器拒绝全屏请求时显示的文案。 */
    fullscreenRejected: string;
    /** 进入全屏失败时生成提示文案的函数。 */
    fullscreenFailed: (reason: string) => string;
  };
  /** 预览器工具栏相关文案。 */
  toolbar: {
    /** 工具栏区域的无障碍名称。 */
    region: string;
    /** 打开文件按钮文案。 */
    selectFile: string;
    /** 上一张幻灯片按钮文案。 */
    previousSlide: string;
    /** 下一张幻灯片按钮文案。 */
    nextSlide: string;
    /** 显示演讲者备注按钮文案。 */
    showSpeakerNotes: string;
    /** 隐藏演讲者备注按钮文案。 */
    hideSpeakerNotes: string;
    /** 演讲者备注标题文案。 */
    speakerNotes: string;
    /** 缩放比例输入框的无障碍名称。 */
    zoomLevel: string;
    /** 缩小预览按钮文案。 */
    zoomOut: string;
    /** 放大预览按钮文案。 */
    zoomIn: string;
    /** 进入全屏按钮文案。 */
    fullscreen: string;
    /** 退出全屏按钮文案。 */
    exitFullscreen: string;
  };
  /** 无可预览内容时显示的文案映射。 */
  empty: Record<PreviewKind, string>;
  /** 加载状态相关文案。 */
  loading: {
    /** 正在解析文件时显示的文案。 */
    parsing: string;
  };
  /** 当前操作产生的错误；未提供表示没有错误。 */
  error: {
    /** 文件预览失败时显示的文案。 */
    previewFailed: string;
  };
  /** 按需内容加载相关文案。 */
  lazyContent: {
    /** 加载状态相关文案。 */
    loading: string;
    /** 重新加载按钮文案。 */
    retry: string;
    /** 页面加载失败时显示的文案。 */
    pageLoadFailed: string;
    /** 幻灯片加载失败时显示的文案。 */
    slideLoadFailed: string;
    /** 工作表加载失败时显示的文案。 */
    sheetLoadFailed: string;
    /** 媒体资源加载失败时显示的文案。 */
    resourceLoadFailed: string;
  };
  /** 文件解析进度相关文案。 */
  progress: {
    /** 各解析阶段对应的显示文案。 */
    stages: Record<ParseStage, string>;
    /** 部分内容已经可预览时显示的标题。 */
    partialTitle: string;
    /** 部分内容已经可预览时显示的补充说明。 */
    partialDescription: string;
  };
  /** Word 大纲相关文案。 */
  outline: {
    /** 大纲区域的无障碍名称。 */
    region: string;
    /** 展开大纲节点操作文案。 */
    expand: string;
    /** 收起操作使用的文案。 */
    collapse: string;
    /** 面向用户展示的标题。 */
    title: string;
    /** 大纲树相关文案。 */
    tree: string;
    /** 调整大纲侧栏宽度时使用的文案。 */
    resize: string;
  };
  /** 电子表格预览相关文案。 */
  spreadsheet: {
    /** 显示模式选择控件的无障碍名称。 */
    viewMode: string;
    /** 保持源文件版式的选项文案。 */
    sourceViewMode: string;
    /** 优先完整显示单元格内容的选项文案。 */
    readingViewMode: string;
    /** 工作表标签区域的无障碍名称。 */
    sheets: string;
    /** 当前元素的尺寸说明。 */
    dimensions: (rows: number, columns: number) => string;
    /** 图片加载失败时显示的文案。 */
    imageLoadFailed: (alt?: string) => string;
  };
  /** 演示文稿预览相关文案。 */
  presentation: {
    /** 幻灯片缩略图导航区域的无障碍名称。 */
    thumbnailsRegion: string;
    /** 当前处理或展示的幻灯片。 */
    slide: (index: number) => string;
    /** 演示文稿包含的幻灯片数量。 */
    slideCount: (count: number) => string;
    /** 演讲者备注区域的无障碍名称。 */
    notesRegion: string;
    /** 调整备注区域高度时使用的文案。 */
    resizeNotes: string;
    /** 没有演讲者备注时显示的文案。 */
    emptyNotes: string;
  };
  /** 内容图片预览、下载和右键菜单相关文案。 */
  imagePreview: {
    /** 图片预览弹层的无障碍名称。 */
    region: string;
    /** 图片右键菜单的无障碍名称。 */
    contextMenu: string;
    /** 为可交互图片生成操作说明。 */
    openLabel: (name?: string) => string;
    /** 打开图片预览操作文案。 */
    preview: string;
    /** 下载原始图片操作文案。 */
    download: string;
    /** 顺时针旋转图片操作文案。 */
    rotate: string;
    /** 恢复图片初始显示状态的操作文案。 */
    reset: string;
    /** 关闭图片预览操作文案。 */
    close: string;
    /** 缩小图片操作文案。 */
    zoomOut: string;
    /** 放大图片操作文案。 */
    zoomIn: string;
    /** 当前图片缩放比例的无障碍名称。 */
    zoomLevel: string;
    /** 图片资源正在加载时显示的文案。 */
    loading: string;
    /** 图片资源加载失败时显示的文案。 */
    loadFailed: string;
    /** 重新加载图片资源的操作文案。 */
    retry: string;
    /** 图片下载失败时生成提示文案的函数。 */
    downloadFailed: (reason: string) => string;
  };
  /** 文档超链接激活、触屏确认和降级相关文案。 */
  hyperlink: {
    /** 超链接右键菜单的无障碍名称。 */
    contextMenu: string;
    /** 打开外部链接的操作文案。 */
    open: string;
    /** 跳转到文档内部目标的操作文案。 */
    jump: string;
    /** 复制外部链接地址的操作文案。 */
    copy: string;
    /** 链接地址复制失败时显示的文案。 */
    copyFailed: string;
    /** 根据当前平台修饰键生成鼠标激活提示。 */
    activationHint: (modifier: 'Ctrl' | 'Command') => string;
    /** 触屏首次点击链接时显示的确认提示。 */
    touchConfirm: string;
    /** 内部链接目标不存在时报告的说明。 */
    targetNotFound: string;
    /** 链接因安全策略被拦截时报告的说明。 */
    blocked: string;
    /** 源格式链接暂时无法可靠解析时报告的说明。 */
    unsupported: string;
  };
  /** 文档图片区域相关文案。 */
  document: {
    /** 图片集标题文案。 */
    images: string;
  };
  /** 图表渲染相关文案。 */
  chart: {
    /** 页面或幻灯片尺寸无效时显示的文案。 */
    invalidSize: string;
    /** 静态图片无法显示时使用的替代文本。 */
    staticAlt: string;
    /** 内容渲染失败时显示的文案。 */
    renderFailed: string;
    /** 地图数据加载失败时显示的文案。 */
    mapLoadFailed: string;
  };
};
