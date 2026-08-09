import type { SpreadsheetViewMode } from '../services/spreadsheet/viewMode';

/** 宿主可以读取或控制的文件预览视图状态。 */
export type OfficeFileViewerViewState = {
  /** 当前缩放比例。 */
  zoom: number;
  /** 当前选中的幻灯片零基索引。 */
  activeSlideIndex: number;
  /** 当前选中的工作表标识。 */
  activeSheetId?: string;
  /** Word 大纲侧栏是否展开。 */
  wordOutlineVisible: boolean;
  /** 演讲者备注区域是否展开。 */
  speakerNotesVisible: boolean;
  /** 电子表格采用的显示模式。 */
  spreadsheetViewMode: SpreadsheetViewMode;
};

/** 单次视图状态变化的字段和值。 */
export type OfficeFileViewerViewStateChange = {
  [StateKey in keyof OfficeFileViewerViewState]-?: {
    /** 本次发生变化的状态字段。 */
    key: StateKey;
    /** 该字段请求更新到的值。 */
    value: OfficeFileViewerViewState[StateKey];
  };
}[keyof OfficeFileViewerViewState];
