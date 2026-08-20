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
  /** 文档全文查找相关文案。 */
  search: {
    /** 查找侧栏区域的无障碍名称。 */
    region: string;
    /** 工具栏和侧栏显示的标题。 */
    title: string;
    /** 展开查找侧栏的操作文案。 */
    expand: string;
    /** 收起查找侧栏的操作文案。 */
    collapse: string;
    /** 搜索输入框占位文案。 */
    placeholder: string;
    /** 切换到上一个匹配结果的操作文案。 */
    previous: string;
    /** 切换到下一个匹配结果的操作文案。 */
    next: string;
    /** 区分大小写选项文案。 */
    matchCase: string;
    /** 全字匹配选项文案。 */
    wholeWord: string;
    /** 尚未输入查询时显示的说明。 */
    emptyQuery: string;
    /** 查询没有匹配结果时显示的说明。 */
    noResults: string;
    /** 扫描尚未完成时显示的说明。 */
    searching: string;
    /** 生成结果数量文案。 */
    resultCount: (count: number) => string;
    /** 生成当前结果位置文案。 */
    currentResult: (current: number, total: number) => string;
    /** 生成 Word 搜索结果的页码文案。 */
    page: (index: number) => string;
    /** 隐藏幻灯片结果的标记文案。 */
    hiddenSlide: string;
    /** 调整查找侧栏宽度时使用的文案。 */
    resize: string;
  };
  /** Office 审阅、批注和修订模式相关文案。 */
  review: {
    /** 审阅面板区域的无障碍名称。 */
    region: string;
    /** 工具栏和面板显示的标题。 */
    title: string;
    /** 展开审阅面板的操作文案。 */
    expand: string;
    /** 收起审阅面板的操作文案。 */
    collapse: string;
    /** 调整审阅面板宽度时使用的文案。 */
    resize: string;
    /** 关闭审阅面板的操作文案。 */
    close: string;
    /** 批注列表的无障碍名称。 */
    comments: string;
    /** 上一条批注操作文案。 */
    previous: string;
    /** 下一条批注操作文案。 */
    next: string;
    /** 没有可列出批注时显示的说明。 */
    empty: string;
    /** 批注正文为空时显示的说明。 */
    emptyComment: string;
    /** 修订没有可见文字摘要时显示的说明。 */
    emptyRevision: string;
    /** 批注作者缺失时显示的说明。 */
    unknownAuthor: string;
    /** 已解决批注的状态文案。 */
    resolved: string;
    /** Word 修订模式控件标题。 */
    revisionMode: string;
    /** Word 最终态修订模式文案。 */
    revisionFinal: string;
    /** Word 标记态修订模式文案。 */
    revisionMarkup: string;
    /** Word 原始态修订模式文案。 */
    revisionOriginal: string;
    /** 插入修订类别文案。 */
    revisionInsert: string;
    /** 删除修订类别文案。 */
    revisionDelete: string;
    /** 移出修订类别文案。 */
    revisionMoveFrom: string;
    /** 移入修订类别文案。 */
    revisionMoveTo: string;
    /** 格式修订类别文案。 */
    revisionFormat: string;
    /** 生成批注数量文案。 */
    itemCount: (count: number) => string;
    /** 生成笔记数量文案。 */
    noteCount: (count: number) => string;
    /** 未知 Word 页面时显示的目标文案。 */
    wordTarget: string;
    /** 生成页面目标文案。 */
    pageTarget: (page: number) => string;
    /** 生成单元格目标文案。 */
    cellTarget: (sheetId: string, row: number, column: number) => string;
    /** 生成幻灯片目标文案。 */
    slideTarget: (slideId: string) => string;
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
    /** 外部媒体未获得加载许可时显示的文案。 */
    externalMediaBlocked: string;
    /** 浏览器无法播放当前媒体时显示的文案。 */
    mediaUnsupported: string;
    /** 媒体资源正在按需加载时显示的文案。 */
    mediaLoading: string;
    /** 下载当前音视频资源的操作文案。 */
    downloadMedia: string;
    /** 为音频或视频控件生成可访问名称。 */
    mediaLabel: (kind: 'audio' | 'video', name?: string) => string;
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
