import type { ParseProgress } from '../../services/parsing';
import type { OfficeFileViewerPreviewState } from '../../services/parsing/internalTypes';
import type { PreviewKind } from '../../services/preview';
import {
  DEFAULT_SPREADSHEET_VIEW_MODE,
  type SpreadsheetViewMode,
} from '../../services/spreadsheet/viewMode';

/** 查看器当前文档所处的加载与交付阶段。 */
export type OfficeViewerDocumentState =
  | { phase: 'empty' }
  | { phase: 'resolving' }
  | {
      phase: 'parsing';
      /** 当前解析文件的名称。 */
      fileName: string;
      /** 当前解析文件的格式。 */
      previewKind: PreviewKind;
      /** 最近一次解析进度。 */
      progress?: ParseProgress;
      /** 当前可用于渐进展示的预览快照。 */
      preview?: OfficeFileViewerPreviewState;
    }
  | {
      phase: 'ready';
      /** 已完成解析的文件名称。 */
      fileName: string;
      /** 已完成交付的预览快照。 */
      preview: OfficeFileViewerPreviewState;
    }
  | {
      phase: 'degraded';
      /** 保留渐进结果的文件名称。 */
      fileName: string;
      /** 解析失败后仍可展示的最后快照。 */
      preview: OfficeFileViewerPreviewState;
    }
  | {
      phase: 'failed';
      /** 已能确定时记录失败文件的名称。 */
      fileName?: string;
      /** 可直接展示给用户的失败说明。 */
      message: string;
    };

/** 与文件内容解耦、但需要随文件切换协调重置的界面状态。 */
export type OfficeViewerViewState = {
  /** 当前选中的幻灯片索引。 */
  activeSlideIndex: number;
  /** 当前选中的工作表标识。 */
  activeSheetId?: string;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 当前实例是否处于全屏状态。 */
  isFullscreen: boolean;
  /** 非受控模式下保存的演讲者备注展开状态。 */
  internalShowSpeakerNotes: boolean;
  /** Word 文档大纲是否展开。 */
  showWordOutline: boolean;
  /** 电子表格当前采用的显示模式。 */
  spreadsheetViewMode: SpreadsheetViewMode;
};

/** 控制器交付给视图层的完整可渲染状态。 */
export type OfficeViewerState = {
  /** 当前文档生命周期。 */
  document: OfficeViewerDocumentState;
  /** 当前界面交互状态。 */
  view: OfficeViewerViewState;
};

/** 查看器状态机支持的全部显式动作。 */
export type OfficeViewerAction =
  | { type: 'resolve-started' }
  | {
      type: 'parse-started';
      /** 本次解析文件的名称。 */
      fileName: string;
      /** 本次解析文件的格式。 */
      previewKind: PreviewKind;
      /** 新文件采用的初始缩放比例。 */
      zoom: number;
      /** 新文件采用的非受控备注初始状态。 */
      showSpeakerNotes: boolean;
    }
  | {
      type: 'progressed';
      /** 最近一次解析进度。 */
      progress: ParseProgress;
    }
  | {
      type: 'partial-installed';
      /** 当前会话最新的渐进预览快照。 */
      preview: OfficeFileViewerPreviewState;
    }
  | {
      type: 'completed';
      /** 已完成解析的文件名称。 */
      fileName: string;
      /** 最终交付的预览快照。 */
      preview: OfficeFileViewerPreviewState;
    }
  | {
      type: 'partial-retained';
      /** 解析失败但保留内容的文件名称。 */
      fileName: string;
      /** 失败后保留的最后预览快照。 */
      preview: OfficeFileViewerPreviewState;
    }
  | {
      type: 'failed';
      /** 已能确定时记录失败文件的名称。 */
      fileName?: string;
      /** 可直接展示给用户的失败说明。 */
      message: string;
    }
  | {
      type: 'slide-selected';
      /** 目标幻灯片索引。 */
      index: number;
    }
  | {
      type: 'sheet-selected';
      /** 目标工作表标识。 */
      sheetId: string;
    }
  | {
      type: 'zoom-changed';
      /** 目标缩放比例。 */
      zoom: number;
    }
  | {
      type: 'fullscreen-changed';
      /** 浏览器确认后的全屏状态。 */
      fullscreen: boolean;
    }
  | {
      type: 'speaker-notes-changed';
      /** 非受控备注面板的目标状态。 */
      visible: boolean;
    }
  | {
      type: 'word-outline-changed';
      /** Word 大纲的目标状态。 */
      visible: boolean;
    }
  | {
      type: 'spreadsheet-view-mode-changed';
      /** 电子表格的目标显示模式。 */
      viewMode: SpreadsheetViewMode;
    };

/** 创建包含调用方默认值的初始查看器状态。 */
export function createInitialOfficeViewerState(options: {
  /** 初始缩放比例。 */
  defaultZoom: number;
  /** 非受控备注面板的初始状态。 */
  defaultShowSpeakerNotes: boolean;
}): OfficeViewerState {
  return {
    document: { phase: 'empty' },
    view: {
      activeSlideIndex: 0,
      activeSheetId: undefined,
      zoom: options.defaultZoom,
      isFullscreen: false,
      internalShowSpeakerNotes: options.defaultShowSpeakerNotes,
      showWordOutline: false,
      spreadsheetViewMode: DEFAULT_SPREADSHEET_VIEW_MODE,
    },
  };
}

/** 返回预览中可用的工作表标识；非表格预览返回 undefined。 */
function getPreviewSheetIds(
  preview: OfficeFileViewerPreviewState,
): string[] | undefined {
  if (preview.mode === 'source') {
    if (preview.previewKind !== 'xls' && preview.previewKind !== 'xlsx') {
      return undefined;
    }
    return preview.summary.sheets.map((sheet) => sheet.id);
  }
  if (preview.model.kind !== 'xls' && preview.model.kind !== 'xlsx') {
    return undefined;
  }
  return preview.model.workbook.sheets.map((sheet) => sheet.id);
}

/** 返回预览中的幻灯片数量；非演示文稿预览返回 undefined。 */
function getPreviewSlideCount(
  preview: OfficeFileViewerPreviewState,
): number | undefined {
  if (preview.mode === 'source') {
    if (preview.previewKind !== 'ppt' && preview.previewKind !== 'pptx') {
      return undefined;
    }
    return preview.summary.slideCount;
  }
  if (preview.model.kind !== 'ppt' && preview.model.kind !== 'pptx') {
    return undefined;
  }
  return preview.model.document.slides.length;
}

/** 根据新快照校正 Sheet 和幻灯片选择，避免保留已不存在的目标。 */
function reconcileViewWithPreview(
  view: OfficeViewerViewState,
  preview: OfficeFileViewerPreviewState,
): OfficeViewerViewState {
  const sheetIds = getPreviewSheetIds(preview);
  const slideCount = getPreviewSlideCount(preview);
  return {
    ...view,
    activeSheetId: sheetIds
      ? view.activeSheetId && sheetIds.includes(view.activeSheetId)
        ? view.activeSheetId
        : sheetIds[0]
      : undefined,
    activeSlideIndex:
      slideCount === undefined
        ? 0
        : Math.min(view.activeSlideIndex, Math.max(0, slideCount - 1)),
  };
}

/** 以可穷尽动作维护文档生命周期和需要同步更新的界面状态。 */
export function officeViewerReducer(
  state: OfficeViewerState,
  action: OfficeViewerAction,
): OfficeViewerState {
  switch (action.type) {
    case 'resolve-started':
      return {
        document: { phase: 'resolving' },
        view: {
          ...state.view,
          activeSlideIndex: 0,
          activeSheetId: undefined,
          showWordOutline: false,
          spreadsheetViewMode: DEFAULT_SPREADSHEET_VIEW_MODE,
        },
      };
    case 'parse-started':
      return {
        document: {
          phase: 'parsing',
          fileName: action.fileName,
          previewKind: action.previewKind,
        },
        view: {
          ...state.view,
          activeSlideIndex: 0,
          activeSheetId: undefined,
          zoom: action.zoom,
          internalShowSpeakerNotes: action.showSpeakerNotes,
          showWordOutline: false,
          spreadsheetViewMode: DEFAULT_SPREADSHEET_VIEW_MODE,
        },
      };
    case 'progressed':
      if (state.document.phase !== 'parsing') return state;
      return {
        ...state,
        document: { ...state.document, progress: action.progress },
      };
    case 'partial-installed':
      if (state.document.phase !== 'parsing') return state;
      return {
        document: {
          ...state.document,
          previewKind: action.preview.previewKind,
          preview: action.preview,
        },
        view: reconcileViewWithPreview(state.view, action.preview),
      };
    case 'completed':
      return {
        document: {
          phase: 'ready',
          fileName: action.fileName,
          preview: action.preview,
        },
        view: reconcileViewWithPreview(state.view, action.preview),
      };
    case 'partial-retained':
      return {
        document: {
          phase: 'degraded',
          fileName: action.fileName,
          preview: action.preview,
        },
        view: reconcileViewWithPreview(state.view, action.preview),
      };
    case 'failed':
      return {
        document: {
          phase: 'failed',
          fileName: action.fileName,
          message: action.message,
        },
        view: {
          ...state.view,
          activeSlideIndex: 0,
          activeSheetId: undefined,
          showWordOutline: false,
          spreadsheetViewMode: DEFAULT_SPREADSHEET_VIEW_MODE,
        },
      };
    case 'slide-selected':
      if (state.view.activeSlideIndex === action.index) return state;
      return {
        ...state,
        view: { ...state.view, activeSlideIndex: action.index },
      };
    case 'sheet-selected':
      if (state.view.activeSheetId === action.sheetId) return state;
      return {
        ...state,
        view: { ...state.view, activeSheetId: action.sheetId },
      };
    case 'zoom-changed':
      if (state.view.zoom === action.zoom) return state;
      return { ...state, view: { ...state.view, zoom: action.zoom } };
    case 'fullscreen-changed':
      if (state.view.isFullscreen === action.fullscreen) return state;
      return {
        ...state,
        view: { ...state.view, isFullscreen: action.fullscreen },
      };
    case 'speaker-notes-changed':
      if (state.view.internalShowSpeakerNotes === action.visible) return state;
      return {
        ...state,
        view: {
          ...state.view,
          internalShowSpeakerNotes: action.visible,
        },
      };
    case 'word-outline-changed':
      if (state.view.showWordOutline === action.visible) return state;
      return {
        ...state,
        view: { ...state.view, showWordOutline: action.visible },
      };
    case 'spreadsheet-view-mode-changed':
      if (state.view.spreadsheetViewMode === action.viewMode) return state;
      return {
        ...state,
        view: { ...state.view, spreadsheetViewMode: action.viewMode },
      };
  }
}
