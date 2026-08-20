/** Word 批注在正文中的稳定字符范围。 */
export type OfficeWordAnnotationTarget = Readonly<{
  kind: 'word-range';
  /** 正文块的稳定标识。 */
  blockId: string;
  /** 批注范围在正文块中的起始字符偏移。 */
  startOffset: number;
  /** 批注范围在正文块中的结束字符偏移。 */
  endOffset: number;
  /** 已知时记录零基页面索引，便于按需页面优先加载。 */
  pageIndex?: number;
}>;

/** Excel 批注对应的工作表单元格。 */
export type OfficeSpreadsheetAnnotationTarget = Readonly<{
  kind: 'spreadsheet-cell';
  /** 工作表稳定标识。 */
  sheetId: string;
  /** 从 1 开始的行索引。 */
  row: number;
  /** 从 1 开始的列索引。 */
  column: number;
}>;

/** PowerPoint 批注对应的幻灯片或具体元素。 */
export type OfficePresentationAnnotationTarget = Readonly<{
  kind: 'presentation-element';
  /** 幻灯片稳定标识。 */
  slideId: string;
  /** 已知时记录批注锚定的元素标识。 */
  elementId?: string;
  /** 已知时记录零基幻灯片索引。 */
  slideIndex?: number;
}>;

/** 不同 Office 格式共用的批注导航目标。 */
export type OfficeAnnotationTarget =
  | OfficeWordAnnotationTarget
  | OfficeSpreadsheetAnnotationTarget
  | OfficePresentationAnnotationTarget;

/** Word 修订内容的只读投影模式。 */
export type WordRevisionMode = 'final' | 'markup' | 'original';

/** 单条 Office 批注及其线程、作者和定位信息。 */
export type OfficeAnnotation = Readonly<{
  /** 在当前文档中稳定且唯一的批注标识。 */
  id: string;
  /** 批注作者；源文件缺失时保持为空。 */
  author?: string;
  /** 源文件提供的 ISO 日期或原始日期文本。 */
  createdAt?: string;
  /** 批注正文纯文本。 */
  text: string;
  /** 批注线程是否已经解决。 */
  resolved?: boolean;
  /** 回复所属的父批注标识。 */
  parentId?: string;
  /** 格式相关的精确导航目标。 */
  target: OfficeAnnotationTarget;
}>;

/** 批注数据源的轻量版本快照。 */
export type OfficeAnnotationSourceSnapshot = Readonly<{
  /** 数据发生变化时单调递增的修订号。 */
  revision: number;
  /** 当前可读取的批注总数。 */
  count: number;
  /** 当前文档包含的修订记录数量。 */
  revisionCount: number;
  /** 当前文档包含的脚注、尾注或其他审阅笔记数量。 */
  noteCount: number;
  /** 当前 Word 文档是否支持最终态、标记态和原始态切换。 */
  supportsRevisionModes: boolean;
}>;

/** OfficeFileViewer 对外开放的只读审阅初始配置。 */
export type OfficeFileViewerReviewOptions = Readonly<{
  /** 非受控模式下是否默认展开审阅面板。 */
  defaultPanelVisible?: boolean;
  /** 非受控模式下 Word 默认采用的修订投影。 */
  defaultRevisionMode?: WordRevisionMode;
  /** 是否显示源文档批注，默认显示。 */
  showComments?: boolean;
  /** 是否显示脚注、尾注和格式笔记，默认显示。 */
  showNotes?: boolean;
}>;
