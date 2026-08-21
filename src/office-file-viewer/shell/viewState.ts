import type { WordRevisionMode } from '../services/annotations/types';
import type { SpreadsheetViewMode } from '../services/spreadsheet/viewMode';

/** 预览器支持的固定比例和自适应缩放模式。 */
export type OfficeFileViewerZoomMode = 'percentage' | 'fit-width' | 'fit-page';
/** 宿主可以读取或控制的文件预览视图状态。 */
export type OfficeFileViewerViewState = {
  /** 当前缩放比例。 */
  zoom: number;
  /** 当前采用固定比例或随视口自适应；旧调用方缺省时按固定比例处理。 */
  zoomMode?: OfficeFileViewerZoomMode;
  /** 当前选中的幻灯片零基索引。 */
  activeSlideIndex: number;
  /** 当前选中的工作表标识。 */
  activeSheetId?: string;
  /** Word 大纲侧栏是否展开。 */
  wordOutlineVisible: boolean;
  /** 文档查找侧栏是否展开。 */
  searchVisible: boolean;
  /** 文档审阅侧栏是否展开。 */
  reviewPanelVisible: boolean;
  /** Word 修订内容采用的只读投影模式。 */
  wordRevisionMode: WordRevisionMode;
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
