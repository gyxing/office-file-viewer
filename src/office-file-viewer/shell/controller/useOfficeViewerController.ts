import type { MutableRefObject, RefObject } from 'react';
import { useCallback, useEffect, useMemo, useReducer, useRef } from 'react';
import type { OfficeFileViewerMessages } from '../../locale';
import {
  ensureSupportedOfficeFile,
  normalizeOfficeFileUri,
  OfficeFileViewerInputError,
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
  createOfficeResourceStore,
  type OfficeResourceStore,
} from '../../services/resource-store';
import {
  createOfficeDocumentSession,
  type OfficeDocumentSession,
} from '../../services/session';
import type { SpreadsheetSource } from '../../services/spreadsheet/SpreadsheetSource';
import type {
  WordOutlineProvider,
  WordOutlineProviderSnapshot,
} from '../../services/word/WordOutlineProvider';
import { useExternalStoreSelector } from '../../shared/react/useExternalStoreSnapshot';
import {
  OFFICE_MAX_ZOOM,
  OFFICE_MIN_ZOOM,
  OFFICE_ZOOM_STEP,
} from '../constants';
import {
  createInitialOfficeViewerState,
  officeViewerReducer,
  type OfficeViewerDocumentState,
  type OfficeViewerState,
} from './officeViewerState';
import { disposeViewerPreviewState } from './previewResources';

/** 保存当前加载代次收到的渐进结果，用于丢弃过期异步输出。 */
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
  onError?: (error: Error, file?: File) => void;
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
  /** 设置指定预览比例。 */
  changeZoom(zoom: number): void;
  /** 恢复调用方当前提供的默认比例。 */
  resetZoom(): void;
  /** 切换演讲者备注面板。 */
  toggleSpeakerNotes(): void;
  /** 切换 Word 大纲侧栏。 */
  toggleWordOutline(): void;
  /** 关闭 Word 大纲侧栏。 */
  closeWordOutline(): void;
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
      defaultZoom: options.defaultZoom,
      defaultShowSpeakerNotes: options.defaultShowSpeakerNotes,
    }),
  );
  const optionsRef = useLatestRef(options);
  const stateRef = useLatestRef(state);
  const viewerRef = useRef<HTMLDivElement | null>(null);
  const resourceStore = useMemo(() => createOfficeResourceStore(), []);
  const resourceStoreEffectGenerationRef = useRef(0);
  const loadGenerationRef = useRef(0);
  const documentSessionIdRef = useRef<string>();
  const requestControllerRef = useRef<AbortController>();
  const parseSessionRef = useRef<OfficeFileViewerParseSession>();
  const previewRef = useRef<OfficeFileViewerPreviewState>();
  const pendingPartialRef = useRef<PendingPartialResult>();
  const partialFrameRef = useRef<number>();
  const preview = getDocumentPreview(state.document);
  const previewKind = getDocumentPreviewKind(state.document);
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
  const presentationSlideCountRef = useLatestRef(presentationSlideCount);
  const hasWordOutlineRef = useLatestRef(hasWordOutline);

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
    async (file: File, loadGeneration: number) => {
      resetActiveSession();
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
        const currentInternalNotes =
          stateRef.current.view.internalShowSpeakerNotes;
        dispatch({
          type: 'parse-started',
          fileName: file.name,
          previewKind: fileKind,
          zoom: currentOptions.defaultZoom,
          showSpeakerNotes:
            currentOptions.showSpeakerNotes === undefined
              ? currentOptions.defaultShowSpeakerNotes
              : currentInternalNotes,
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
        const normalizedError =
          nextError instanceof Error
            ? nextError
            : new Error(optionsRef.current.messages.file.parseFailed);
        if (!retainedPartial) {
          if (documentSessionIdRef.current === documentSession?.id) {
            documentSessionIdRef.current = undefined;
          }
          dispatch({
            type: 'failed',
            fileName: documentSession ? file.name : undefined,
            message:
              normalizedError instanceof OfficeFileViewerInputError
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
      requestControllerRef.current?.abort();
      requestControllerRef.current = undefined;
      const loadGeneration = ++loadGenerationRef.current;
      await loadFile(file, loadGeneration);
    },
    [loadFile],
  );

  useEffect(() => {
    if (!options.uri) return;

    resetActiveSession();
    dispatch({ type: 'resolve-started' });
    // 固化本次 effect 的文件来源，避免异步闭包丢失类型收窄。
    const uriToLoad = options.uri;
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
        file = await normalizeOfficeFileUri(
          uriToLoad,
          optionsRef.current.messages,
          requestController?.signal,
        );
        if (loadGeneration !== loadGenerationRef.current) return;
        await loadFile(file, loadGeneration);
      } catch (nextError) {
        if (
          loadGeneration !== loadGenerationRef.current ||
          requestController?.signal.aborted
        ) {
          return;
        }

        const normalizedError =
          nextError instanceof Error
            ? nextError
            : new Error(optionsRef.current.messages.file.loadFailed);
        dispatch({
          type: 'failed',
          message:
            normalizedError instanceof OfficeFileViewerInputError
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
    return () => {
      requestController?.abort();
    };
  }, [loadFile, options.uri, resetActiveSession]);

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

  const previousSlide = useCallback(() => {
    dispatch({
      type: 'slide-selected',
      index: Math.max(stateRef.current.view.activeSlideIndex - 1, 0),
    });
  }, [stateRef]);

  const nextSlide = useCallback(() => {
    dispatch({
      type: 'slide-selected',
      index: Math.min(
        stateRef.current.view.activeSlideIndex + 1,
        Math.max(0, presentationSlideCountRef.current - 1),
      ),
    });
  }, [presentationSlideCountRef, stateRef]);

  const selectSlide = useCallback((index: number) => {
    dispatch({ type: 'slide-selected', index });
  }, []);

  const selectSheet = useCallback((sheetId: string) => {
    dispatch({ type: 'sheet-selected', sheetId });
  }, []);

  const zoomOut = useCallback(() => {
    dispatch({
      type: 'zoom-changed',
      zoom: Math.max(
        OFFICE_MIN_ZOOM,
        stateRef.current.view.zoom - OFFICE_ZOOM_STEP,
      ),
    });
  }, [stateRef]);

  const zoomIn = useCallback(() => {
    dispatch({
      type: 'zoom-changed',
      zoom: Math.min(
        OFFICE_MAX_ZOOM,
        stateRef.current.view.zoom + OFFICE_ZOOM_STEP,
      ),
    });
  }, [stateRef]);

  const changeZoom = useCallback((zoom: number) => {
    dispatch({ type: 'zoom-changed', zoom });
  }, []);

  const resetZoom = useCallback(() => {
    dispatch({
      type: 'zoom-changed',
      zoom: optionsRef.current.defaultZoom,
    });
  }, [optionsRef]);

  const toggleSpeakerNotes = useCallback(() => {
    const currentOptions = optionsRef.current;
    const nextVisible = !(
      currentOptions.showSpeakerNotes ??
      stateRef.current.view.internalShowSpeakerNotes
    );
    if (currentOptions.showSpeakerNotes === undefined) {
      dispatch({ type: 'speaker-notes-changed', visible: nextVisible });
    }
    currentOptions.onSpeakerNotesVisibilityChange?.(nextVisible);
  }, [optionsRef, stateRef]);

  const toggleWordOutline = useCallback(() => {
    if (!hasWordOutlineRef.current) return;
    dispatch({
      type: 'word-outline-changed',
      visible: !stateRef.current.view.showWordOutline,
    });
  }, [hasWordOutlineRef, stateRef]);

  const closeWordOutline = useCallback(() => {
    dispatch({ type: 'word-outline-changed', visible: false });
  }, []);

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
        new Error(optionsRef.current.messages.file.fullscreenFailed(reason)),
      );
    }
  }, [optionsRef]);

  const speakerNotesVisible =
    options.showSpeakerNotes ?? state.view.internalShowSpeakerNotes;
  const fullscreenSupported =
    typeof document !== 'undefined' &&
    typeof document.documentElement.requestFullscreen === 'function';
  const canGoPreviousSlide =
    previewFamily === 'presentation' &&
    presentationSlideCount > 1 &&
    state.view.activeSlideIndex > 0;
  const canGoNextSlide =
    previewFamily === 'presentation' &&
    presentationSlideCount > 1 &&
    state.view.activeSlideIndex < presentationSlideCount - 1;
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
      previousSlide,
      nextSlide,
      selectSlide,
      selectSheet,
      zoomOut,
      zoomIn,
      changeZoom,
      resetZoom,
      toggleSpeakerNotes,
      toggleWordOutline,
      closeWordOutline,
      toggleFullscreen,
    }),
    [
      changeZoom,
      closeWordOutline,
      nextSlide,
      previousSlide,
      resetZoom,
      selectFile,
      selectSheet,
      selectSlide,
      toggleFullscreen,
      toggleSpeakerNotes,
      toggleWordOutline,
      zoomIn,
      zoomOut,
    ],
  );
  const meta = useMemo<OfficeViewerMeta>(
    () => ({
      preview,
      previewKind,
      format,
      hasRenderableContent,
      speakerNotesVisible,
      fullscreenSupported,
      loadingTip,
    }),
    [
      format,
      fullscreenSupported,
      hasRenderableContent,
      loadingTip,
      preview,
      previewKind,
      speakerNotesVisible,
    ],
  );

  return { state, actions, meta, viewerRef, resourceStore };
}
