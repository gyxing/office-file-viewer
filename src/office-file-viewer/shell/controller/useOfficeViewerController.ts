import type { MutableRefObject, RefObject } from 'react';
import { useCallback, useEffect, useMemo, useReducer, useRef } from 'react';
import type { OfficeFileViewerMessages } from '../../locale';
import type { WordRevisionMode } from '../../services/annotations/types';
import {
  normalizeOfficeFileViewerError,
  OfficeFileViewerError,
} from '../../services/errors/OfficeFileViewerError';
import {
  ensureSupportedOfficeFile,
  normalizeOfficeFileUri,
  type OfficeFileViewerUri,
} from '../../services/input/normalizeOfficeFileUri';
import type {
  OfficeParseOptions,
  OfficePreviewReadyInfo,
  ParseProgress,
} from '../../services/parsing';
import { createOfficeFileViewerParseSession } from '../../services/parsing/createParseSession';
import type {
  OfficeFileViewerParseSession,
  OfficeFileViewerPreviewHandle,
  OfficeFileViewerPreviewState,
} from '../../services/parsing/internalTypes';
import type { PresentationSource } from '../../services/presentation/PresentationSource';
import {
  detectPreviewKind,
  getPreviewFamily,
  isPresentationPreviewKind,
  isSpreadsheetPreviewKind,
  type ParsedOfficeFile,
  type PreviewKind,
} from '../../services/preview';
import {
  collectOfficeFileWarnings,
  collectOfficePreviewWarnings,
  type OfficeFileViewerWarning,
} from '../../services/previewWarnings';
import {
  createOfficeResourceStore,
  type OfficeResourceStore,
} from '../../services/resource-store';
import {
  createOfficeDocumentSession,
  type OfficeDocumentSession,
} from '../../services/session';
import type { SpreadsheetSource } from '../../services/spreadsheet/SpreadsheetSource';
import {
  DEFAULT_SPREADSHEET_VIEW_MODE,
  type SpreadsheetViewMode,
} from '../../services/spreadsheet/viewMode';
import type {
  WordOutlineProvider,
  WordOutlineProviderSnapshot,
} from '../../services/word/WordOutlineProvider';
import { useExternalStoreSelector } from '../../shared/react/useExternalStoreSnapshot';
import { OFFICE_ZOOM_STEP } from '../constants';
import { normalizeOfficeZoom } from '../normalizeOfficeZoom';
import type {
  OfficeFileViewerViewState,
  OfficeFileViewerViewStateChange,
  OfficeFileViewerZoomMode,
} from '../viewState';
import {
  createInitialOfficeViewerState,
  officeViewerReducer,
  type OfficeViewerDocumentState,
  type OfficeViewerState,
} from './officeViewerState';
import { disposeViewerPreviewState } from './previewResources';

/** 保存当前加载代次收到的渐进结果，用于丢弃过期异步输出。 */
type OfficeViewerLoadSource =
  | { kind: 'file'; file: File }
  | { kind: 'uri'; uri: OfficeFileViewerUri };

type PendingPartialResult = {
  /** 本次加载的递增代次。 */
  loadGeneration: number;
  /** 当前加载代次产生的预览快照。 */
  preview: OfficeFileViewerPreviewState;
};

/** 控制器依赖的输入、默认值和公共回调。 */
export type UseOfficeViewerControllerOptions = {
  /** 待预览文件的来源。 */
  uri?: OfficeFileViewerUri;
  /** 新文件和重置操作使用的缩放比例。 */
  defaultZoom: number;
  /** 当前实例是否启用文档搜索能力。 */
  searchEnabled: boolean;
  /** 非受控搜索侧栏的默认状态。 */
  defaultSearchVisible: boolean;
  /** 当前实例是否启用文档审阅能力。 */
  reviewEnabled: boolean;
  /** 非受控审阅侧栏的默认状态。 */
  defaultReviewPanelVisible: boolean;
  /** 非受控 Word 修订内容的默认投影模式。 */
  defaultWordRevisionMode: WordRevisionMode;
  /** 非受控模式下各视图字段的统一初始值。 */
  defaultViewState?: Partial<OfficeFileViewerViewState>;
  /** 由宿主按字段控制的视图状态。 */
  viewState?: Partial<OfficeFileViewerViewState>;
  /** 用户请求改变视图状态时触发。 */
  onViewStateChange?: (
    state: OfficeFileViewerViewState,
    change: OfficeFileViewerViewStateChange,
  ) => void;
  /** 非受控备注面板的默认状态。 */
  defaultShowSpeakerNotes: boolean;
  /** 受控模式下的备注面板状态。 */
  showSpeakerNotes?: boolean;
  /** 备注面板状态变化回调。 */
  onSpeakerNotesVisibilityChange?: (visible: boolean) => void;
  /** 完整物化文件解析成功后的回调。 */
  onFileParsed?: (parsed: ParsedOfficeFile, file: File) => void;
  /** 首屏预览可以使用后的回调。 */
  onPreviewReady?: (info: OfficePreviewReadyInfo, file: File) => void;
  /** 文件加载、解析或全屏失败后的回调。 */
  onError?: (error: OfficeFileViewerError, file?: File) => void;
  /** 解析降级或格式兼容警告回调。 */
  onWarning?: (warning: OfficeFileViewerWarning, file: File) => void;
  /** 底层解析会话配置。 */
  parseOptions?: OfficeParseOptions;
  /** 解析进度变化回调。 */
  onParseProgress?: (progress: ParseProgress) => void;
  /** 当前语言环境对应的界面文案。 */
  messages: OfficeFileViewerMessages;
};

/** 视图层可以触发的稳定用户动作集合。 */
export type OfficeViewerActions = {
  /** 解析用户在文件选择器中选中的文件。 */
  selectFile(file: File): Promise<void>;
  /** 重新加载最近一次文件来源。 */
  retry(): void;
  /** 切换到上一张幻灯片。 */
  previousSlide(): void;
  /** 切换到下一张幻灯片。 */
  nextSlide(): void;
  /** 选择指定幻灯片。 */
  selectSlide(index: number): void;
  /** 选择指定工作表。 */
  selectSheet(sheetId: string): void;
  /** 减小预览比例。 */
  zoomOut(): void;
  /** 放大预览比例。 */
  zoomIn(): void;
  /** 设置指定预览比例，并切换回固定比例模式。 */
  changeZoom(zoom: number): void;
  /** 切换固定比例或自适应缩放模式。 */
  changeZoomMode(mode: OfficeFileViewerZoomMode): void;
  /** 自适应测量完成后更新比例，但保持当前缩放模式。 */
  applyFitZoom(zoom: number): void;
  /** 设置电子表格的显示模式。 */
  changeSpreadsheetViewMode(viewMode: SpreadsheetViewMode): void;
  /** 切换演讲者备注面板。 */
  toggleSpeakerNotes(): void;
  /** 切换 Word 大纲侧栏。 */
  toggleWordOutline(): void;
  /** 关闭 Word 大纲侧栏。 */
  closeWordOutline(): void;
  /** 切换文档查找侧栏。 */
  toggleSearch(): void;
  /** 打开文档查找侧栏。 */
  openSearch(): void;
  /** 关闭文档查找侧栏。 */
  closeSearch(): void;
  /** 切换文档审阅侧栏。 */
  toggleReviewPanel(): void;
  /** 打开文档审阅侧栏。 */
  openReviewPanel(): void;
  /** 关闭文档审阅侧栏。 */
  closeReviewPanel(): void;
  /** 设置 Word 修订内容的投影模式。 */
  changeWordRevisionMode(mode: WordRevisionMode): void;
  /** 进入或退出浏览器全屏。 */
  toggleFullscreen(): Promise<void>;
};

/** 当前格式可向组合层公开的只读能力。 */
export type OfficeViewerFormatMeta =
  | { kind: 'empty' }
  | {
      kind: 'presentation';
      /** 当前演示文稿的幻灯片数量。 */
      slideCount: number;
      /** 当前页是否允许向前切换。 */
      canPrevious: boolean;
      /** 当前页是否允许向后切换。 */
      canNext: boolean;
    }
  | {
      kind: 'word';
      /** 当前文字文档是否提供可导航的大纲。 */
      hasOutline: boolean;
    }
  | { kind: 'spreadsheet' };

/** 从文档状态和外部 Store 推导出的只读渲染能力。 */
export type OfficeViewerMeta = {
  /** 当前可展示或最终交付的预览快照。 */
  preview?: OfficeFileViewerPreviewState;
  /** 当前文件识别出的格式。 */
  previewKind?: PreviewKind;
  /** 当前加载代次正在预览的文件。 */
  currentFile?: File;
  /** 可用于解析文档相对链接的 HTTP(S) 来源地址。 */
  sourceUrl?: string;
  /** 当前格式对应的专属渲染能力。 */
  format: OfficeViewerFormatMeta;
  /** 当前是否已有可以交给格式查看器的内容。 */
  hasRenderableContent: boolean;
  /** 受控或非受控模式合并后的备注面板状态。 */
  speakerNotesVisible: boolean;
  /** 当前浏览器是否支持元素全屏。 */
  fullscreenSupported: boolean;
  /** 当前解析阶段对应的本地化提示。 */
  loadingTip?: string;
  /** 当前错误是否存在可重新加载的文件来源。 */
  canRetry: boolean;
};

/** 控制器向公共组件交付的状态、动作和实例资源。 */
export type OfficeViewerControllerResult = {
  /** 查看器当前可渲染状态。 */
  state: OfficeViewerState;
  /** 用户交互动作。 */
  actions: OfficeViewerActions;
  /** 不允许视图层反向修改的推导能力。 */
  meta: OfficeViewerMeta;
  /** 全屏操作绑定的查看器根元素。 */
  viewerRef: RefObject<HTMLDivElement>;
  /** 当前查看器实例独占的资源存储。 */
  resourceStore: OfficeResourceStore;
};

/** 保存最新值，同时让异步生命周期函数保持稳定引用。 */
function useLatestRef<Value>(value: Value): MutableRefObject<Value> {
  const valueRef = useRef(value);
  valueRef.current = value;
  return valueRef;
}

/** 调用外部观察回调；异常异步上报，避免反向污染查看器状态事务。 */
function notifyObserver<TArguments extends unknown[]>(
  observer: ((...args: TArguments) => void) | undefined,
  ...args: TArguments
) {
  if (!observer) return;
  try {
    observer(...args);
  } catch (observerError) {
    setTimeout(() => {
      throw observerError;
    }, 0);
  }
}

/** 将幻灯片索引规范为当前演示文稿可导航的范围。 */
function normalizeSlideIndex(value: number, slideCount: number) {
  const fallback = 0;
  if (!Number.isFinite(value) || slideCount <= 0) return fallback;
  return Math.min(Math.max(0, Math.trunc(value)), slideCount - 1);
}

/** 忽略 JavaScript 调用方传入的无效自适应缩放模式。 */
function normalizeOfficeZoomMode(
  value: OfficeFileViewerZoomMode | undefined,
): OfficeFileViewerZoomMode {
  return value === 'fit-width' || value === 'fit-page' ? value : 'percentage';
}
/** 忽略 JavaScript 调用方传入的无效电子表格显示模式。 */
function normalizeSpreadsheetViewMode(
  value: SpreadsheetViewMode | undefined,
): SpreadsheetViewMode {
  return value === 'reading' || value === 'source'
    ? value
    : DEFAULT_SPREADSHEET_VIEW_MODE;
}

/** 忽略 JavaScript 调用方传入的无效 Word 修订模式。 */
function normalizeWordRevisionMode(
  value: WordRevisionMode | undefined,
): WordRevisionMode {
  return value === 'markup' || value === 'original' ? value : 'final';
}

/** 合并旧版默认属性和统一默认视图状态，保持已有调用方式兼容。 */
function createDefaultViewState(
  options: UseOfficeViewerControllerOptions,
): OfficeFileViewerViewState {
  const defaults = options.defaultViewState;
  return {
    zoom: normalizeOfficeZoom(defaults?.zoom ?? options.defaultZoom),
    zoomMode: normalizeOfficeZoomMode(defaults?.zoomMode),
    activeSlideIndex: normalizeSlideIndex(
      defaults?.activeSlideIndex ?? 0,
      Number.MAX_SAFE_INTEGER,
    ),
    activeSheetId: defaults?.activeSheetId,
    wordOutlineVisible: defaults?.wordOutlineVisible ?? false,
    searchVisible:
      options.searchEnabled &&
      (defaults?.searchVisible ?? options.defaultSearchVisible),
    reviewPanelVisible:
      options.reviewEnabled &&
      (defaults?.reviewPanelVisible ?? options.defaultReviewPanelVisible),
    wordRevisionMode: normalizeWordRevisionMode(
      defaults?.wordRevisionMode ?? options.defaultWordRevisionMode,
    ),
    speakerNotesVisible:
      defaults?.speakerNotesVisible ?? options.defaultShowSpeakerNotes,
    spreadsheetViewMode: normalizeSpreadsheetViewMode(
      defaults?.spreadsheetViewMode,
    ),
  };
}

/** 将内部字段转换为公开视图状态，屏蔽全屏和内部备注存储细节。 */
function createPublicViewState(
  view: OfficeViewerState['view'],
): OfficeFileViewerViewState {
  return {
    zoom: view.zoom,
    zoomMode: view.zoomMode,
    activeSlideIndex: view.activeSlideIndex,
    activeSheetId: view.activeSheetId,
    wordOutlineVisible: view.showWordOutline,
    searchVisible: view.showSearch,
    reviewPanelVisible: view.showReviewPanel,
    wordRevisionMode: view.wordRevisionMode,
    speakerNotesVisible: view.internalShowSpeakerNotes,
    spreadsheetViewMode: view.spreadsheetViewMode,
  };
}

/** 把单字段变更合并到完整公开状态，保留字段和值的联合类型关系。 */
function applyPublicViewStateChange(
  state: OfficeFileViewerViewState,
  change: OfficeFileViewerViewStateChange,
): OfficeFileViewerViewState {
  switch (change.key) {
    case 'zoom':
      return { ...state, zoom: change.value };
    case 'zoomMode':
      return { ...state, zoomMode: change.value };
    case 'activeSlideIndex':
      return { ...state, activeSlideIndex: change.value };
    case 'activeSheetId':
      return { ...state, activeSheetId: change.value };
    case 'wordOutlineVisible':
      return { ...state, wordOutlineVisible: change.value };
    case 'searchVisible':
      return { ...state, searchVisible: change.value };
    case 'reviewPanelVisible':
      return { ...state, reviewPanelVisible: change.value };
    case 'wordRevisionMode':
      return { ...state, wordRevisionMode: change.value };
    case 'speakerNotesVisible':
      return { ...state, speakerNotesVisible: change.value };
    case 'spreadsheetViewMode':
      return { ...state, spreadsheetViewMode: change.value };
  }
}
/** 读取当前电子表格预览中有效的工作表标识。 */
function getSpreadsheetSheetIds(
  preview: OfficeFileViewerPreviewState | undefined,
): string[] {
  if (!preview || !isSpreadsheetPreviewKind(preview.previewKind)) return [];
  if (preview.mode === 'source') {
    return preview.source.getSnapshot().sheets.map((sheet) => sheet.id);
  }
  return preview.model.workbook.sheets.map((sheet) => sheet.id);
}
/** 从生命周期状态中读取当前可展示的预览快照。 */
function getDocumentPreview(
  documentState: OfficeViewerDocumentState,
): OfficeFileViewerPreviewState | undefined {
  if (documentState.phase === 'parsing') return documentState.preview;
  if (documentState.phase === 'ready' || documentState.phase === 'degraded') {
    return documentState.preview;
  }
  return undefined;
}

/** 从生命周期状态中读取已识别的文件格式。 */
function getDocumentPreviewKind(
  documentState: OfficeViewerDocumentState,
): PreviewKind | undefined {
  if (documentState.phase === 'parsing') return documentState.previewKind;
  return getDocumentPreview(documentState)?.previewKind;
}

/** 提取当前演示文稿按需数据源。 */
function getPresentationSource(
  preview: OfficeFileViewerPreviewState | undefined,
): PresentationSource | undefined {
  if (
    preview?.mode === 'source' &&
    isPresentationPreviewKind(preview.previewKind)
  ) {
    return preview.source;
  }
  return undefined;
}

/** 提取当前电子表格按需数据源。 */
function getSpreadsheetSource(
  preview: OfficeFileViewerPreviewState | undefined,
): SpreadsheetSource | undefined {
  if (
    preview?.mode === 'source' &&
    isSpreadsheetPreviewKind(preview.previewKind)
  ) {
    return preview.source;
  }
  return undefined;
}

/** 提取当前 Word 按需数据源的大纲 Provider。 */
function getWordOutlineProvider(
  preview: OfficeFileViewerPreviewState | undefined,
): WordOutlineProvider | undefined {
  if (
    preview?.mode === 'source' &&
    (preview.previewKind === 'doc' || preview.previewKind === 'docx')
  ) {
    return preview.source.outline;
  }
  return undefined;
}

/** 计算物化 Word 模型中的大纲数量。 */
function getMaterializedWordOutlineCount(
  preview: OfficeFileViewerPreviewState | undefined,
): number {
  if (preview?.mode !== 'materialized') return 0;
  if (preview.model.kind === 'doc' || preview.model.kind === 'docx') {
    return preview.model.document.outline?.length ?? 0;
  }
  return 0;
}

/** 只读取控制器导航能力依赖的幻灯片数量。 */
function selectPresentationSlideCount(
  snapshot: ReturnType<PresentationSource['getSnapshot']>,
) {
  return snapshot.slideCount;
}

/** 只读取控制器空状态判断依赖的工作表数量。 */
function selectSpreadsheetSheetCount(
  snapshot: ReturnType<SpreadsheetSource['getSnapshot']>,
) {
  return snapshot.sheets.length;
}

/** 只读取工具栏显隐依赖的大纲数量。 */
function selectWordOutlineCount(snapshot: WordOutlineProviderSnapshot) {
  return snapshot.count;
}

/** 计算演示文稿当前可导航的幻灯片数量。 */
function getPresentationSlideCount(
  preview: OfficeFileViewerPreviewState | undefined,
  sourceSlideCount: number,
): number {
  if (!preview) return 0;
  if (
    preview.mode === 'source' &&
    isPresentationPreviewKind(preview.previewKind)
  ) {
    return sourceSlideCount;
  }
  if (
    preview.mode === 'materialized' &&
    isPresentationPreviewKind(preview.model.kind)
  ) {
    return preview.model.document.slides.length;
  }
  return 0;
}

/** 判断当前快照是否已有格式查看器可以渲染的内容。 */
function getHasRenderableContent(
  preview: OfficeFileViewerPreviewState | undefined,
  presentationSlideCount: number,
  spreadsheetSheetCount: number,
): boolean {
  if (!preview) return false;
  if (preview.mode === 'source') {
    if (isPresentationPreviewKind(preview.previewKind)) {
      return presentationSlideCount > 0;
    }
    if (isSpreadsheetPreviewKind(preview.previewKind)) {
      return spreadsheetSheetCount > 0;
    }
    if (preview.previewKind === 'docx') return true;
    return preview.source.getSnapshot().pages.length > 0;
  }
  if (isPresentationPreviewKind(preview.model.kind)) {
    return preview.model.document.slides.length > 0;
  }
  if (isSpreadsheetPreviewKind(preview.model.kind)) {
    return preview.model.workbook.sheets.length > 0;
  }
  if (preview.model.kind === 'docx') {
    return preview.model.document.blocks.length > 0;
  }
  return preview.model.document.paragraphs.length > 0;
}

/** 管理文件输入、解析会话、资源所有权和查看器交互状态。 */
export function useOfficeViewerController(
  options: UseOfficeViewerControllerOptions,
): OfficeViewerControllerResult {
  const [state, dispatch] = useReducer(
    officeViewerReducer,
    createInitialOfficeViewerState({
      defaultViewState: createDefaultViewState(options),
    }),
  );
  const optionsRef = useLatestRef(options);
  const viewerRef = useRef<HTMLDivElement | null>(null);
  const resourceStore = useMemo(() => createOfficeResourceStore(), []);
  const resourceStoreEffectGenerationRef = useRef(0);
  const loadGenerationRef = useRef(0);
  const lastLoadSourceRef = useRef<OfficeViewerLoadSource>();
  const documentSessionIdRef = useRef<string>();
  const currentFileRef = useRef<File>();
  const sourceUrlRef = useRef<string>();
  const requestControllerRef = useRef<AbortController>();
  const parseSessionRef = useRef<OfficeFileViewerParseSession>();
  const previewRef = useRef<OfficeFileViewerPreviewState>();
  const pendingPartialRef = useRef<PendingPartialResult>();
  const partialFrameRef = useRef<number>();
  const preview = getDocumentPreview(state.document);
  const previewKind = getDocumentPreviewKind(state.document);
  const currentFile = currentFileRef.current;
  const sourceUrl = sourceUrlRef.current;
  const previewFamily = previewKind ? getPreviewFamily(previewKind) : undefined;
  const presentationSource = getPresentationSource(preview);
  const spreadsheetSource = getSpreadsheetSource(preview);
  const wordOutlineProvider = getWordOutlineProvider(preview);
  const sourcePresentationSlideCount = useExternalStoreSelector(
    presentationSource,
    selectPresentationSlideCount,
    0,
  );
  const sourceSpreadsheetSheetCount = useExternalStoreSelector(
    spreadsheetSource,
    selectSpreadsheetSheetCount,
    0,
  );
  const sourceWordOutlineCount = useExternalStoreSelector(
    wordOutlineProvider,
    selectWordOutlineCount,
    0,
  );
  const presentationSlideCount = getPresentationSlideCount(
    preview,
    sourcePresentationSlideCount,
  );
  const hasRenderableContent = getHasRenderableContent(
    preview,
    sourcePresentationSlideCount,
    sourceSpreadsheetSheetCount,
  );
  const hasWordOutline =
    (sourceWordOutlineCount || getMaterializedWordOutlineCount(preview)) > 0;
  const spreadsheetSheetIds = getSpreadsheetSheetIds(preview);
  const controlledViewState = options.viewState;
  const controlledSheetId = controlledViewState?.activeSheetId;
  const internalSheetId = state.view.activeSheetId;
  const activeSheetId =
    controlledSheetId && spreadsheetSheetIds.includes(controlledSheetId)
      ? controlledSheetId
      : internalSheetId && spreadsheetSheetIds.includes(internalSheetId)
      ? internalSheetId
      : spreadsheetSheetIds[0];
  const speakerNotesVisible =
    controlledViewState?.speakerNotesVisible ??
    options.showSpeakerNotes ??
    state.view.internalShowSpeakerNotes;
  const effectiveViewState: OfficeViewerState['view'] = {
    ...state.view,
    activeSlideIndex: normalizeSlideIndex(
      controlledViewState?.activeSlideIndex ?? state.view.activeSlideIndex,
      presentationSlideCount,
    ),
    activeSheetId,
    zoom: normalizeOfficeZoom(
      controlledViewState?.zoom ?? state.view.zoom,
      state.view.zoom,
    ),
    zoomMode: normalizeOfficeZoomMode(
      controlledViewState?.zoomMode ?? state.view.zoomMode,
    ),
    internalShowSpeakerNotes: speakerNotesVisible,
    showSearch:
      options.searchEnabled &&
      (controlledViewState?.searchVisible ?? state.view.showSearch),
    showReviewPanel:
      options.reviewEnabled &&
      (controlledViewState?.reviewPanelVisible ?? state.view.showReviewPanel),
    wordRevisionMode: normalizeWordRevisionMode(
      controlledViewState?.wordRevisionMode ?? state.view.wordRevisionMode,
    ),
    showWordOutline:
      hasWordOutline &&
      (controlledViewState?.wordOutlineVisible ?? state.view.showWordOutline) &&
      !(
        options.searchEnabled &&
        (controlledViewState?.searchVisible ?? state.view.showSearch)
      ),
    spreadsheetViewMode: normalizeSpreadsheetViewMode(
      controlledViewState?.spreadsheetViewMode ??
        state.view.spreadsheetViewMode,
    ),
  };
  const presentationSlideCountRef = useLatestRef(presentationSlideCount);
  const spreadsheetSheetIdsRef = useLatestRef(spreadsheetSheetIds);
  const hasWordOutlineRef = useLatestRef(hasWordOutline);
  const effectiveViewStateRef = useLatestRef(effectiveViewState);

  const notifyViewStateChange = useCallback(
    (change: OfficeFileViewerViewStateChange) => {
      const nextState = applyPublicViewStateChange(
        createPublicViewState(effectiveViewStateRef.current),
        change,
      );
      notifyObserver(optionsRef.current.onViewStateChange, nextState, change);
    },
    [effectiveViewStateRef, optionsRef],
  );

  const cancelPartialFrame = useCallback(() => {
    if (
      partialFrameRef.current !== undefined &&
      typeof window !== 'undefined'
    ) {
      window.cancelAnimationFrame(partialFrameRef.current);
    }
    partialFrameRef.current = undefined;
    pendingPartialRef.current = undefined;
  }, []);

  const disposeCurrentPreview = useCallback(() => {
    const ownedPreview = previewRef.current;
    previewRef.current = undefined;
    void disposeViewerPreviewState(ownedPreview);
  }, []);

  const resetActiveSession = useCallback(() => {
    parseSessionRef.current?.cancel();
    parseSessionRef.current?.dispose();
    parseSessionRef.current = undefined;
    documentSessionIdRef.current = undefined;
    currentFileRef.current = undefined;
    sourceUrlRef.current = undefined;
    cancelPartialFrame();
    disposeCurrentPreview();
  }, [cancelPartialFrame, disposeCurrentPreview]);

  const installPartialSnapshot = useCallback(
    (nextPreview: OfficeFileViewerPreviewState) => {
      // 同一 sessionId 的快照共享底层资源，替换 partial 时只更新最新所有者。
      previewRef.current = nextPreview;
      dispatch({ type: 'partial-installed', preview: nextPreview });
    },
    [],
  );

  const schedulePartialSnapshot = useCallback(
    (nextPreview: OfficeFileViewerPreviewState, loadGeneration: number) => {
      pendingPartialRef.current = {
        preview: nextPreview,
        loadGeneration,
      };
      if (partialFrameRef.current !== undefined) return;
      if (typeof window === 'undefined') {
        if (
          loadGeneration === loadGenerationRef.current &&
          nextPreview.sessionId === documentSessionIdRef.current
        ) {
          installPartialSnapshot(nextPreview);
        }
        pendingPartialRef.current = undefined;
        return;
      }
      partialFrameRef.current = window.requestAnimationFrame(() => {
        partialFrameRef.current = undefined;
        const pending = pendingPartialRef.current;
        pendingPartialRef.current = undefined;
        if (
          pending &&
          pending.loadGeneration === loadGenerationRef.current &&
          pending.preview.sessionId === documentSessionIdRef.current
        ) {
          installPartialSnapshot(pending.preview);
        }
      });
    },
    [installPartialSnapshot],
  );

  const loadFile = useCallback(
    async (file: File, loadGeneration: number, nextSourceUrl?: string) => {
      resetActiveSession();
      currentFileRef.current = file;
      sourceUrlRef.current = nextSourceUrl;
      dispatch({ type: 'resolve-started' });
      let retainedPartial = false;
      let documentSession: OfficeDocumentSession | undefined;
      let parsedModel: ParsedOfficeFile | undefined;
      let previewReadyInfo: OfficePreviewReadyInfo | undefined;

      try {
        ensureSupportedOfficeFile(file, optionsRef.current.messages);
        if (loadGeneration !== loadGenerationRef.current) return;

        const fileKind = detectPreviewKind(file.name);
        const currentOptions = optionsRef.current;
        dispatch({
          type: 'parse-started',
          fileName: file.name,
          previewKind: fileKind,
          viewState: createDefaultViewState(currentOptions),
        });

        documentSession = createOfficeDocumentSession();
        documentSessionIdRef.current = documentSession.id;
        const isCurrentSession = () =>
          loadGeneration === loadGenerationRef.current &&
          documentSession?.id === documentSessionIdRef.current;
        const parseSession = createOfficeFileViewerParseSession(
          file,
          currentOptions.parseOptions,
          documentSession,
        );
        parseSessionRef.current = parseSession;
        const unsubscribeProgress = parseSession.subscribe((progress) => {
          if (!isCurrentSession()) return;
          dispatch({ type: 'progressed', progress });
          optionsRef.current.onParseProgress?.(progress);
        });
        const unsubscribePartial = parseSession.subscribePartial(
          (nextPreview) => {
            if (
              !isCurrentSession() ||
              nextPreview.sessionId !== documentSession?.id
            ) {
              return;
            }
            schedulePartialSnapshot(nextPreview, loadGeneration);
          },
        );
        let finalPreview: OfficeFileViewerPreviewHandle;
        try {
          finalPreview = await parseSession.result;
        } catch (nextError) {
          cancelPartialFrame();
          const partial = parseSession.partialResult;
          if (partial) {
            if (
              !isCurrentSession() ||
              partial.sessionId !== documentSession.id
            ) {
              await disposeViewerPreviewState(partial);
            } else {
              previewRef.current = partial;
              dispatch({
                type: 'partial-retained',
                fileName: file.name,
                preview: partial,
              });
              notifyObserver(
                optionsRef.current.onWarning,
                {
                  code: 'PARTIAL_PREVIEW_RETAINED',
                  message:
                    optionsRef.current.messages.progress.partialDescription,
                  previewKind: partial.previewKind,
                  source: 'partial-preview',
                },
                file,
              );
              retainedPartial = true;
            }
          }
          throw nextError;
        } finally {
          unsubscribeProgress();
          unsubscribePartial();
          if (parseSessionRef.current === parseSession) {
            parseSession.dispose();
            parseSessionRef.current = undefined;
          }
        }

        if (
          !isCurrentSession() ||
          finalPreview.sessionId !== documentSession.id
        ) {
          await disposeViewerPreviewState(finalPreview);
          return;
        }

        cancelPartialFrame();
        previewRef.current = finalPreview;
        dispatch({
          type: 'completed',
          fileName: file.name,
          preview: finalPreview,
        });
        previewReadyInfo = {
          previewKind: finalPreview.previewKind,
          mode: finalPreview.mode,
        };
        [
          ...collectOfficeFileWarnings(file.name, finalPreview.previewKind),
          ...collectOfficePreviewWarnings(finalPreview),
        ].forEach((warning) => {
          notifyObserver(optionsRef.current.onWarning, warning, file);
        });
        if (finalPreview.mode === 'materialized') {
          parsedModel = finalPreview.model;
        }
      } catch (nextError) {
        if (
          loadGeneration !== loadGenerationRef.current ||
          (documentSession &&
            documentSession.id !== documentSessionIdRef.current)
        ) {
          return;
        }

        // 界面只展示可本地化的概述，原始错误仍交给调用方诊断。
        const normalizedError = normalizeOfficeFileViewerError(nextError, {
          stage: 'parsing',
          fileName: file.name,
          previewKind: documentSession
            ? detectPreviewKind(file.name)
            : undefined,
        });
        if (!retainedPartial) {
          if (documentSessionIdRef.current === documentSession?.id) {
            documentSessionIdRef.current = undefined;
          }
          dispatch({
            type: 'failed',
            fileName: documentSession ? file.name : undefined,
            message:
              normalizedError.code !== 'PARSE_FAILED'
                ? normalizedError.message
                : optionsRef.current.messages.file.parseFailed,
          });
        }
        notifyObserver(optionsRef.current.onError, normalizedError, file);
        return;
      }

      if (parsedModel) {
        notifyObserver(optionsRef.current.onFileParsed, parsedModel, file);
      }
      if (previewReadyInfo) {
        notifyObserver(
          optionsRef.current.onPreviewReady,
          previewReadyInfo,
          file,
        );
      }
    },
    [cancelPartialFrame, resetActiveSession, schedulePartialSnapshot],
  );

  const selectFile = useCallback(
    async (file: File) => {
      lastLoadSourceRef.current = { kind: 'file', file };
      requestControllerRef.current?.abort();
      requestControllerRef.current = undefined;
      const loadGeneration = ++loadGenerationRef.current;
      await loadFile(file, loadGeneration);
    },
    [loadFile],
  );

  const beginUriLoad = useCallback(
    (uriToLoad: OfficeFileViewerUri) => {
      resetActiveSession();
      dispatch({ type: 'resolve-started' });
      const loadGeneration = ++loadGenerationRef.current;
      const requestController =
        typeof AbortController === 'undefined'
          ? undefined
          : new AbortController();
      requestControllerRef.current?.abort();
      requestControllerRef.current = requestController;

      async function loadUri() {
        let file: File | undefined;
        try {
          const normalized = await normalizeOfficeFileUri(
            uriToLoad,
            optionsRef.current.messages,
            requestController?.signal,
          );
          file = normalized.file;
          if (loadGeneration !== loadGenerationRef.current) return;
          await loadFile(file, loadGeneration, normalized.sourceUrl);
        } catch (nextError) {
          if (
            loadGeneration !== loadGenerationRef.current ||
            requestController?.signal.aborted
          ) {
            return;
          }

          const normalizedError = normalizeOfficeFileViewerError(nextError, {
            stage: 'input',
            fileName: file?.name,
          });
          dispatch({
            type: 'failed',
            message:
              normalizedError.code !== 'PARSE_FAILED'
                ? normalizedError.message
                : optionsRef.current.messages.file.loadFailed,
          });
          notifyObserver(optionsRef.current.onError, normalizedError, file);
        } finally {
          if (requestControllerRef.current === requestController) {
            requestControllerRef.current = undefined;
          }
        }
      }

      void loadUri();
      return () => requestController?.abort();
    },
    [loadFile, resetActiveSession],
  );

  useEffect(() => {
    if (!options.uri) return;
    lastLoadSourceRef.current = { kind: 'uri', uri: options.uri };
    return beginUriLoad(options.uri);
  }, [beginUriLoad, options.uri]);

  const retry = useCallback(() => {
    const source = lastLoadSourceRef.current;
    if (!source) return;
    if (source.kind === 'file') {
      void selectFile(source.file);
      return;
    }
    beginUriLoad(source.uri);
  }, [beginUriLoad, selectFile]);

  useEffect(() => {
    const effectGeneration = ++resourceStoreEffectGenerationRef.current;
    return () => {
      // 使所有不可取消的本地解析任务失效，避免卸载后继续回写状态。
      loadGenerationRef.current += 1;
      requestControllerRef.current?.abort();
      requestControllerRef.current = undefined;
      resetActiveSession();
      // StrictMode 会重放 effect；延迟到微任务确认没有后续 setup 再释放 Store。
      void Promise.resolve().then(() => {
        if (resourceStoreEffectGenerationRef.current === effectGeneration) {
          return resourceStore.dispose();
        }
        return undefined;
      });
    };
  }, [resetActiveSession, resourceStore]);

  useEffect(() => {
    if (!hasWordOutline) {
      dispatch({ type: 'word-outline-changed', visible: false });
    }
  }, [hasWordOutline]);

  useEffect(() => {
    if (typeof document === 'undefined') return;

    // 浏览器和 ESC 键都可能改变全屏状态，因此以 fullscreenchange 为唯一来源。
    const handleFullscreenChange = () => {
      dispatch({
        type: 'fullscreen-changed',
        fullscreen: document.fullscreenElement === viewerRef.current,
      });
    };
    document.addEventListener('fullscreenchange', handleFullscreenChange);
    return () => {
      document.removeEventListener('fullscreenchange', handleFullscreenChange);
    };
  }, []);

  const commitSlideSelection = useCallback(
    (index: number) => {
      const normalizedIndex = normalizeSlideIndex(
        index,
        presentationSlideCountRef.current,
      );
      if (effectiveViewStateRef.current.activeSlideIndex === normalizedIndex) {
        return;
      }
      if (optionsRef.current.viewState?.activeSlideIndex === undefined) {
        dispatch({ type: 'slide-selected', index: normalizedIndex });
      }
      notifyViewStateChange({
        key: 'activeSlideIndex',
        value: normalizedIndex,
      });
    },
    [effectiveViewStateRef, notifyViewStateChange, optionsRef],
  );

  const previousSlide = useCallback(() => {
    commitSlideSelection(effectiveViewStateRef.current.activeSlideIndex - 1);
  }, [commitSlideSelection, effectiveViewStateRef]);

  const nextSlide = useCallback(() => {
    commitSlideSelection(effectiveViewStateRef.current.activeSlideIndex + 1);
  }, [commitSlideSelection, effectiveViewStateRef]);

  const selectSlide = useCallback(
    (index: number) => commitSlideSelection(index),
    [commitSlideSelection],
  );

  const selectSheet = useCallback(
    (sheetId: string) => {
      if (!spreadsheetSheetIdsRef.current.includes(sheetId)) return;
      if (effectiveViewStateRef.current.activeSheetId === sheetId) return;
      if (optionsRef.current.viewState?.activeSheetId === undefined) {
        dispatch({ type: 'sheet-selected', sheetId });
      }
      notifyViewStateChange({ key: 'activeSheetId', value: sheetId });
    },
    [
      effectiveViewStateRef,
      notifyViewStateChange,
      optionsRef,
      spreadsheetSheetIdsRef,
    ],
  );

  const commitManualZoom = useCallback(
    (zoom: number) => {
      const currentView = effectiveViewStateRef.current;
      const normalizedZoom = normalizeOfficeZoom(zoom, currentView.zoom);
      const modeChanged = currentView.zoomMode !== 'percentage';
      const zoomChanged = currentView.zoom !== normalizedZoom;
      if (!modeChanged && !zoomChanged) return;

      if (modeChanged && optionsRef.current.viewState?.zoomMode === undefined) {
        dispatch({ type: 'zoom-mode-changed', mode: 'percentage' });
      }
      if (zoomChanged && optionsRef.current.viewState?.zoom === undefined) {
        dispatch({ type: 'zoom-changed', zoom: normalizedZoom });
      }
      const nextState = {
        ...createPublicViewState(currentView),
        zoom: normalizedZoom,
        zoomMode: 'percentage' as const,
      };
      if (modeChanged) {
        notifyObserver(optionsRef.current.onViewStateChange, nextState, {
          key: 'zoomMode',
          value: 'percentage',
        });
      }
      if (zoomChanged) {
        notifyObserver(optionsRef.current.onViewStateChange, nextState, {
          key: 'zoom',
          value: normalizedZoom,
        });
      }
    },
    [effectiveViewStateRef, optionsRef],
  );

  const applyFitZoom = useCallback(
    (zoom: number) => {
      const currentView = effectiveViewStateRef.current;
      const normalizedZoom = normalizeOfficeZoom(zoom, currentView.zoom);
      if (currentView.zoom === normalizedZoom) return;
      if (optionsRef.current.viewState?.zoom === undefined) {
        dispatch({ type: 'zoom-changed', zoom: normalizedZoom });
      }
      notifyObserver(
        optionsRef.current.onViewStateChange,
        {
          ...createPublicViewState(currentView),
          zoom: normalizedZoom,
        },
        { key: 'zoom', value: normalizedZoom },
      );
    },
    [effectiveViewStateRef, optionsRef],
  );

  const changeZoomMode = useCallback(
    (mode: OfficeFileViewerZoomMode) => {
      const normalizedMode = normalizeOfficeZoomMode(mode);
      const currentView = effectiveViewStateRef.current;
      if (currentView.zoomMode === normalizedMode) return;
      if (optionsRef.current.viewState?.zoomMode === undefined) {
        dispatch({ type: 'zoom-mode-changed', mode: normalizedMode });
      }
      notifyObserver(
        optionsRef.current.onViewStateChange,
        {
          ...createPublicViewState(currentView),
          zoomMode: normalizedMode,
        },
        { key: 'zoomMode', value: normalizedMode },
      );
    },
    [effectiveViewStateRef, optionsRef],
  );

  const zoomOut = useCallback(() => {
    commitManualZoom(effectiveViewStateRef.current.zoom - OFFICE_ZOOM_STEP);
  }, [commitManualZoom, effectiveViewStateRef]);

  const zoomIn = useCallback(() => {
    commitManualZoom(effectiveViewStateRef.current.zoom + OFFICE_ZOOM_STEP);
  }, [commitManualZoom, effectiveViewStateRef]);

  const changeZoom = useCallback(
    (zoom: number) => commitManualZoom(zoom),
    [commitManualZoom],
  );

  const changeSpreadsheetViewMode = useCallback(
    (viewMode: SpreadsheetViewMode) => {
      const normalizedMode = normalizeSpreadsheetViewMode(viewMode);
      if (
        effectiveViewStateRef.current.spreadsheetViewMode === normalizedMode
      ) {
        return;
      }
      if (optionsRef.current.viewState?.spreadsheetViewMode === undefined) {
        dispatch({
          type: 'spreadsheet-view-mode-changed',
          viewMode: normalizedMode,
        });
      }
      notifyViewStateChange({
        key: 'spreadsheetViewMode',
        value: normalizedMode,
      });
    },
    [effectiveViewStateRef, notifyViewStateChange, optionsRef],
  );

  const toggleSpeakerNotes = useCallback(() => {
    const currentOptions = optionsRef.current;
    const nextVisible = !effectiveViewStateRef.current.internalShowSpeakerNotes;
    const controlled =
      currentOptions.viewState?.speakerNotesVisible !== undefined ||
      currentOptions.showSpeakerNotes !== undefined;
    if (!controlled) {
      dispatch({ type: 'speaker-notes-changed', visible: nextVisible });
    }
    notifyObserver(currentOptions.onSpeakerNotesVisibilityChange, nextVisible);
    notifyViewStateChange({
      key: 'speakerNotesVisible',
      value: nextVisible,
    });
  }, [effectiveViewStateRef, notifyViewStateChange, optionsRef]);

  const commitSearchVisibility = useCallback(
    (visible: boolean) => {
      const nextVisible = visible && optionsRef.current.searchEnabled;
      if (effectiveViewStateRef.current.showSearch === nextVisible) return;
      if (nextVisible && effectiveViewStateRef.current.showWordOutline) {
        if (optionsRef.current.viewState?.wordOutlineVisible === undefined) {
          dispatch({ type: 'word-outline-changed', visible: false });
        }
        notifyViewStateChange({ key: 'wordOutlineVisible', value: false });
      }
      if (optionsRef.current.viewState?.searchVisible === undefined) {
        dispatch({ type: 'search-changed', visible: nextVisible });
      }
      notifyViewStateChange({ key: 'searchVisible', value: nextVisible });
    },
    [effectiveViewStateRef, notifyViewStateChange, optionsRef],
  );

  const toggleSearch = useCallback(() => {
    commitSearchVisibility(!effectiveViewStateRef.current.showSearch);
  }, [commitSearchVisibility, effectiveViewStateRef]);

  const openSearch = useCallback(() => {
    commitSearchVisibility(true);
  }, [commitSearchVisibility]);

  const closeSearch = useCallback(() => {
    commitSearchVisibility(false);
  }, [commitSearchVisibility]);

  const commitWordOutlineVisibility = useCallback(
    (visible: boolean) => {
      const nextVisible = visible && hasWordOutlineRef.current;
      if (effectiveViewStateRef.current.showWordOutline === nextVisible) return;
      if (nextVisible && effectiveViewStateRef.current.showSearch) {
        commitSearchVisibility(false);
      }
      if (optionsRef.current.viewState?.wordOutlineVisible === undefined) {
        dispatch({ type: 'word-outline-changed', visible: nextVisible });
      }
      notifyViewStateChange({
        key: 'wordOutlineVisible',
        value: nextVisible,
      });
    },
    [
      effectiveViewStateRef,
      hasWordOutlineRef,
      commitSearchVisibility,
      notifyViewStateChange,
      optionsRef,
    ],
  );

  const toggleWordOutline = useCallback(() => {
    commitWordOutlineVisibility(!effectiveViewStateRef.current.showWordOutline);
  }, [commitWordOutlineVisibility, effectiveViewStateRef]);

  const closeWordOutline = useCallback(() => {
    commitWordOutlineVisibility(false);
  }, [commitWordOutlineVisibility]);

  const commitReviewPanelVisibility = useCallback(
    (visible: boolean) => {
      const nextVisible = visible && optionsRef.current.reviewEnabled;
      if (effectiveViewStateRef.current.showReviewPanel === nextVisible) return;
      if (optionsRef.current.viewState?.reviewPanelVisible === undefined) {
        dispatch({ type: 'review-panel-changed', visible: nextVisible });
      }
      notifyViewStateChange({ key: 'reviewPanelVisible', value: nextVisible });
    },
    [effectiveViewStateRef, notifyViewStateChange, optionsRef],
  );

  const toggleReviewPanel = useCallback(() => {
    commitReviewPanelVisibility(!effectiveViewStateRef.current.showReviewPanel);
  }, [commitReviewPanelVisibility, effectiveViewStateRef]);

  const openReviewPanel = useCallback(() => {
    commitReviewPanelVisibility(true);
  }, [commitReviewPanelVisibility]);

  const closeReviewPanel = useCallback(() => {
    commitReviewPanelVisibility(false);
  }, [commitReviewPanelVisibility]);

  const changeWordRevisionMode = useCallback(
    (mode: WordRevisionMode) => {
      const nextMode = normalizeWordRevisionMode(mode);
      if (effectiveViewStateRef.current.wordRevisionMode === nextMode) return;
      if (optionsRef.current.viewState?.wordRevisionMode === undefined) {
        dispatch({ type: 'word-revision-mode-changed', mode: nextMode });
      }
      notifyViewStateChange({ key: 'wordRevisionMode', value: nextMode });
    },
    [effectiveViewStateRef, notifyViewStateChange, optionsRef],
  );

  const toggleFullscreen = useCallback(async () => {
    const viewer = viewerRef.current;
    if (
      !viewer ||
      typeof document === 'undefined' ||
      typeof viewer.requestFullscreen !== 'function'
    ) {
      return;
    }

    try {
      if (document.fullscreenElement === viewer) {
        await document.exitFullscreen();
      } else {
        await viewer.requestFullscreen();
      }
    } catch (nextError) {
      const reason =
        nextError instanceof Error
          ? nextError.message
          : optionsRef.current.messages.file.fullscreenRejected;
      notifyObserver(
        optionsRef.current.onError,
        new OfficeFileViewerError(
          'FULLSCREEN_FAILED',
          optionsRef.current.messages.file.fullscreenFailed(reason),
          { stage: 'fullscreen', recoverable: true, cause: nextError },
        ),
      );
    }
  }, [optionsRef]);

  const fullscreenSupported =
    typeof document !== 'undefined' &&
    typeof document.documentElement.requestFullscreen === 'function';
  const canGoPreviousSlide =
    previewFamily === 'presentation' &&
    presentationSlideCount > 1 &&
    effectiveViewState.activeSlideIndex > 0;
  const canGoNextSlide =
    previewFamily === 'presentation' &&
    presentationSlideCount > 1 &&
    effectiveViewState.activeSlideIndex < presentationSlideCount - 1;
  const loadingTip =
    state.document.phase === 'parsing' && state.document.progress
      ? options.messages.progress.stages[state.document.progress.stage]
      : undefined;
  const format = useMemo<OfficeViewerFormatMeta>(() => {
    if (previewFamily === 'presentation') {
      return {
        kind: 'presentation',
        slideCount: presentationSlideCount,
        canPrevious: canGoPreviousSlide,
        canNext: canGoNextSlide,
      };
    }
    if (previewFamily === 'word') {
      return { kind: 'word', hasOutline: hasWordOutline };
    }
    if (previewFamily === 'spreadsheet') {
      return { kind: 'spreadsheet' };
    }
    return { kind: 'empty' };
  }, [
    canGoNextSlide,
    canGoPreviousSlide,
    hasWordOutline,
    presentationSlideCount,
    previewFamily,
  ]);

  const actions = useMemo<OfficeViewerActions>(
    () => ({
      selectFile,
      retry,
      previousSlide,
      nextSlide,
      selectSlide,
      selectSheet,
      zoomOut,
      zoomIn,
      changeZoom,
      changeZoomMode,
      applyFitZoom,
      changeSpreadsheetViewMode,
      toggleSpeakerNotes,
      toggleWordOutline,
      closeWordOutline,
      toggleSearch,
      openSearch,
      closeSearch,
      toggleReviewPanel,
      openReviewPanel,
      closeReviewPanel,
      changeWordRevisionMode,
      toggleFullscreen,
    }),
    [
      changeZoom,
      changeZoomMode,
      applyFitZoom,
      changeSpreadsheetViewMode,
      changeWordRevisionMode,
      closeReviewPanel,
      closeSearch,
      closeWordOutline,
      nextSlide,
      previousSlide,
      openReviewPanel,
      openSearch,
      retry,
      selectFile,
      selectSheet,
      selectSlide,
      toggleFullscreen,
      toggleReviewPanel,
      toggleSpeakerNotes,
      toggleSearch,
      toggleWordOutline,
      zoomIn,
      zoomOut,
    ],
  );
  const meta = useMemo<OfficeViewerMeta>(
    () => ({
      preview,
      previewKind,
      currentFile,
      sourceUrl,
      format,
      hasRenderableContent,
      speakerNotesVisible,
      fullscreenSupported,
      loadingTip,
      canRetry: Boolean(lastLoadSourceRef.current),
    }),
    [
      format,
      fullscreenSupported,
      hasRenderableContent,
      loadingTip,
      preview,
      previewKind,
      currentFile,
      sourceUrl,
      speakerNotesVisible,
    ],
  );

  return {
    state: { ...state, view: effectiveViewState },
    actions,
    meta,
    viewerRef,
    resourceStore,
  };
}
