/** 搜索请求中稳定且可序列化的匹配选项。 */
export type OfficeSearchQuery = Readonly<{
  /** 用户输入的原始搜索文本。 */
  text: string;
  /** 是否区分英文字母大小写。 */
  matchCase: boolean;
  /** 是否只匹配完整的英文或数字单词。 */
  wholeWord: boolean;
}>;

/** OfficeFileViewer 对外开放的搜索初始配置。 */
export type OfficeFileViewerSearchOptions = Readonly<{
  /** 非受控模式下是否默认展开查找侧栏。 */
  defaultVisible?: boolean;
  /** 初始是否区分英文字母大小写。 */
  matchCase?: boolean;
  /** 初始是否只匹配完整的英文或数字单词。 */
  wholeWord?: boolean;
}>;

/** Word 搜索结果对应的正文块和原始字符区间。 */
export type OfficeWordSearchTarget = Readonly<{
  kind: 'word';
  /** 正文块的稳定标识。 */
  blockId: string;
  /** 已知时记录零基页面索引；渐进分页期间可以暂时缺省。 */
  pageIndex?: number;
  /** 匹配内容在正文块原始文本中的起始偏移。 */
  startOffset: number;
  /** 匹配内容在正文块原始文本中的结束偏移。 */
  endOffset: number;
}>;

/** Excel 搜索结果对应的工作表单元格。 */
export type OfficeSpreadsheetSearchTarget = Readonly<{
  kind: 'spreadsheet';
  /** 工作表稳定标识。 */
  sheetId: string;
  /** 从 1 开始的行索引。 */
  rowIndex: number;
  /** 从 1 开始的列索引。 */
  columnIndex: number;
}>;

/** PowerPoint 搜索结果对应的幻灯片元素和字符区间。 */
export type OfficePresentationSearchTarget = Readonly<{
  kind: 'presentation';
  /** 零基幻灯片索引。 */
  slideIndex: number;
  /** 幻灯片元素的稳定标识。 */
  elementId: string;
  /** 匹配内容在元素可见文本中的起始偏移。 */
  startOffset: number;
  /** 匹配内容在元素可见文本中的结束偏移。 */
  endOffset: number;
  /** 当前结果是否位于隐藏幻灯片。 */
  hidden: boolean;
}>;

/** 不同 Office 格式使用的精确结果定位信息。 */
export type OfficeSearchTarget =
  | OfficeWordSearchTarget
  | OfficeSpreadsheetSearchTarget
  | OfficePresentationSearchTarget;

/** 单个搜索匹配及其列表展示内容。 */
export type OfficeSearchResult = Readonly<{
  /** 在当前文档和查询中可复现的稳定标识。 */
  id: string;
  /** 与查询命中的原始显示文本。 */
  matchText: string;
  /** 结果列表用于提供上下文的短文本。 */
  previewText: string;
  /** 格式相关的精确导航目标。 */
  target: OfficeSearchTarget;
}>;

/** 搜索提供器向会话报告的增量进度。 */
export type OfficeSearchProgress = Readonly<{
  /** 本批新增的全部结果，不限制文档总结果数。 */
  items: readonly OfficeSearchResult[];
  /** 当前查询已经扫描的内容单元数量。 */
  scanned: number;
  /** 当前已知或最终确认的内容单元总数。 */
  total: number;
  /** 是否已经完成当前查询。 */
  complete: boolean;
}>;

/** 带查询版本的搜索批次，用于阻止旧查询结果回写。 */
export type OfficeSearchBatch = OfficeSearchProgress &
  Readonly<{
    /** 当前会话内单调递增的查询标识。 */
    queryId: number;
  }>;

/** 搜索提供器的结果输出函数。 */
export type OfficeSearchProgressEmitter = (
  progress: OfficeSearchProgress,
) => void;

/** 搜索会话的结果输出函数。 */
export type OfficeSearchBatchEmitter = (batch: OfficeSearchBatch) => void;

/** 格式解析器提供的可取消增量搜索能力。 */
export interface OfficeSearchProvider {
  /** 提供器所对应的预览格式家族。 */
  readonly kind: OfficeSearchTarget['kind'];
  /** 扫描当前文档并按批次输出全部匹配结果。 */
  search(
    query: OfficeSearchQuery,
    emit: OfficeSearchProgressEmitter,
    signal: AbortSignal,
  ): Promise<void>;
}

/** 管理查询版本、取消和旧批次隔离的搜索会话。 */
export interface OfficeSearchSession {
  /** 取消上一查询并执行新查询，完成后返回本次查询标识。 */
  search(
    query: OfficeSearchQuery,
    emit: OfficeSearchBatchEmitter,
  ): Promise<number>;
  /** 取消当前尚未完成的查询。 */
  cancel(): void;
  /** 释放会话并永久阻止后续查询。 */
  dispose(): void;
}
