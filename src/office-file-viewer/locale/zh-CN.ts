import type { OfficeFileViewerMessages } from './types';

/** OfficeFileViewer 默认使用的简体中文界面文案。 */
export const zhCN: OfficeFileViewerMessages = {
  file: {
    unloaded: '未加载文件',
    selectToPreview: '请先选择文件开始预览',
    unsupported:
      '暂不支持该文件类型，请选择 PPTX、PPT、XLSX、XLS、DOCX、DOC 或 WPS 文件',
    unrecognized:
      '无法识别 Office 文件类型，请提供 PPTX、PPT、XLSX、XLS、DOCX、DOC 或 WPS 文件',
    invalidUri:
      'uri 必须是 File、URL 字符串，或返回 File、Blob、URL、Response 的异步函数',
    parseFailed: '文件解析失败',
    loadFailed: '文件加载失败',
    downloadFailed: (status, statusText) =>
      `文件下载失败：${status} ${statusText}`,
    fullscreenRejected: '浏览器拒绝了全屏请求',
    fullscreenFailed: (reason) => `全屏操作失败：${reason}`,
  },
  toolbar: {
    selectFile: '打开文件',
    previousSlide: '上一页',
    nextSlide: '下一页',
    showSpeakerNotes: '显示演讲者备注',
    hideSpeakerNotes: '隐藏演讲者备注',
    speakerNotes: '备注',
    zoomOut: '缩小',
    zoomIn: '放大',
    fullscreen: '全屏',
    exitFullscreen: '退出全屏',
  },
  empty: {
    pptx: '请先上传 PPTX 文件开始预览',
    ppt: '请先上传 PPT 文件开始预览',
    xlsx: '请先上传 XLSX 文件开始预览',
    xls: '请先上传 XLS 文件开始预览',
    docx: '请先上传 DOCX 文件开始预览',
    doc: '请先上传 DOC 或 WPS 文件开始预览',
  },
  loading: { parsing: '正在解析文件' },
  error: { previewFailed: '预览失败' },
  lazyContent: {
    loading: '正在加载内容',
    retry: '重试',
    pageLoadFailed: '页面加载失败',
    slideLoadFailed: '幻灯片加载失败',
    sheetLoadFailed: '工作表加载失败',
    resourceLoadFailed: '文档资源加载失败',
  },
  progress: {
    stages: {
      reading: '正在读取文件',
      container: '正在打开文件容器',
      structure: '正在读取文档结构',
      content: '正在解析文档内容',
      resources: '正在处理文档资源',
      assembling: '正在生成预览',
    },
    partialTitle: '文档解析未完成',
    partialDescription: '当前仅展示已成功解析的部分内容。',
  },
  outline: {
    region: '文档大纲',
    expand: '展开文档大纲',
    collapse: '收起文档大纲',
    title: '大纲',
    tree: '大纲目录',
    resize: '调整文档大纲宽度',
  },
  spreadsheet: {
    dimensions: (rows, columns) => `${rows} 行 × ${columns} 列`,
    imageLoadFailed: (alt) => (alt ? `${alt}（图片加载失败）` : '图片加载失败'),
  },
  presentation: {
    slide: (index) => `第 ${index} 页`,
    slideCount: (count) => `共 ${count} 页`,
    notesRegion: '演讲者备注',
    resizeNotes: '调整演讲者备注高度',
    emptyNotes: '本页无演讲者备注',
  },
  document: { images: '文档图片' },
  chart: {
    invalidSize: '图表尺寸无效',
    staticAlt: '静态图表',
    renderFailed: '图表渲染失败',
    mapLoadFailed: '地图数据加载失败',
  },
};
