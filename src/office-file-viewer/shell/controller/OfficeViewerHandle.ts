import type { MutableRefObject, RefObject } from 'react';
import type {
  OfficeCapabilities,
  OfficeDocumentRuntime,
} from '../../core/types';
import type { WordPageNavigationController } from '../../formats/word-pages/types';
import { OfficeFileViewerError } from '../../services/errors/OfficeFileViewerError';
import type { OfficeFileViewerViewState } from '../viewState';
import { getOfficeViewerCapabilities } from './officeViewerCapabilities';
import type { OfficeViewerState } from './officeViewerState';
import type {
  OfficeViewerActions,
  OfficeViewerMeta,
} from './useOfficeViewerController';

/** OfficeFileViewer 对外暴露的同步控制句柄。 */
export type OfficeViewerHandle = {
  openFile(file: File): void;
  reload(): void;
  setZoom(value: number): void;
  zoomIn(): void;
  zoomOut(): void;
  fitWidth(): void;
  fitPage(): void;
  goToPage(index: number): void;
  goToSheet(id: string): void;
  goToSlide(index: number): void;
  toggleFullscreen(): void;
  openSearch(): void;
  openReview(): void;
  getViewState(): OfficeFileViewerViewState;
  getCapabilities(): OfficeCapabilities | undefined;
};

export type OfficeViewerHandleRefs = {
  actions: MutableRefObject<OfficeViewerActions>;
  meta: MutableRefObject<OfficeViewerMeta>;
  state: MutableRefObject<OfficeViewerState>;
  customRuntime: MutableRefObject<OfficeDocumentRuntime | undefined>;
  /** 查看器根节点，用于非虚拟分页模式的同步页面定位。 */
  viewerRef?: RefObject<HTMLDivElement>;
  /** 组件卸载后置为 false，保证保留在宿主手中的句柄不再触发动作。 */
  active?: MutableRefObject<boolean>;
  pageNavigationController?: MutableRefObject<
    WordPageNavigationController | undefined
  >;
  onNavigationError?: MutableRefObject<
    ((error: OfficeFileViewerError) => void) | undefined
  >;
};

function isHandleActive(refs: OfficeViewerHandleRefs) {
  return refs.active?.current !== false;
}

function hasRenderableDocument(refs: OfficeViewerHandleRefs) {
  return isHandleActive(refs) && refs.meta.current.hasRenderableContent;
}

function getCurrentCapabilities(refs: OfficeViewerHandleRefs) {
  return getOfficeViewerCapabilities(
    refs.meta.current,
    refs.customRuntime.current,
  );
}

function supportsDocumentAction(
  refs: OfficeViewerHandleRefs,
  capability: keyof Pick<
    OfficeCapabilities,
    | 'canSearch'
    | 'canReview'
    | 'canNavigatePages'
    | 'canNavigateSheets'
    | 'canNavigateSlides'
    | 'canFitPage'
    | 'canFitWidth'
    | 'canFullscreen'
  >,
) {
  if (!hasRenderableDocument(refs)) return false;
  return Boolean(getCurrentCapabilities(refs)?.[capability]);
}

function findStaticWordPages(viewer: HTMLDivElement | null) {
  if (!viewer) return [] as HTMLElement[];
  if (viewer.querySelector('.office-file-word-pages')) return [];
  return Array.from(
    viewer.querySelectorAll<HTMLElement>(
      '.office-file-doc-page-frame, .office-file-docx-page-frame',
    ),
  );
}

/** 创建稳定句柄，所有读取和动作都从最新 Ref 获取。 */
export function createOfficeViewerHandle(
  refs: OfficeViewerHandleRefs,
): OfficeViewerHandle {
  let pendingPageIndex: number | undefined;
  let pendingPageAttempts = 0;
  let pendingPageTimer: ReturnType<typeof setTimeout> | undefined;

  const clearPendingPageNavigation = () => {
    if (pendingPageTimer !== undefined) clearTimeout(pendingPageTimer);
    pendingPageTimer = undefined;
    pendingPageIndex = undefined;
    pendingPageAttempts = 0;
  };

  const reportNavigationError = (error: OfficeFileViewerError) => {
    refs.onNavigationError?.current?.(error);
  };

  const navigatePage = (index: number, canQueue: boolean) => {
    if (!isHandleActive(refs)) {
      clearPendingPageNavigation();
      return;
    }
    const controller = refs.pageNavigationController?.current;
    if (!controller) {
      // 小文档采用完整页面列表，不会挂载 VirtualWordPageList；直接使用页面
      // 框架完成定位，避免句柄把可用的静态分页误判为不支持导航。
      const staticPages = findStaticWordPages(refs.viewerRef?.current ?? null);
      if (staticPages.length) {
        clearPendingPageNavigation();
        if (index >= staticPages.length) {
          throw new OfficeFileViewerError(
            'NAVIGATION_OUT_OF_RANGE',
            `页面索引超出范围：${index}`,
            { stage: 'navigation', recoverable: false },
          );
        }
        const target = staticPages[index];
        if (typeof target.scrollIntoView === 'function') {
          target.scrollIntoView({ block: 'start', behavior: 'auto' });
        }
        return;
      }
      if (
        canQueue &&
        (refs.state.current.document.phase === 'parsing' ||
          refs.state.current.document.phase === 'ready') &&
        typeof setTimeout === 'function'
      ) {
        if (pendingPageIndex !== index) pendingPageAttempts = 0;
        pendingPageIndex = index;
        pendingPageAttempts += 1;
        if (pendingPageAttempts <= 120) {
          pendingPageTimer = setTimeout(() => {
            pendingPageTimer = undefined;
            if (pendingPageIndex === undefined) return;
            navigatePage(pendingPageIndex, true);
          }, 16);
          return;
        }
        clearPendingPageNavigation();
        reportNavigationError(
          new OfficeFileViewerError(
            'NAVIGATION_UNSUPPORTED',
            '当前文档页面导航尚未就绪',
            { stage: 'navigation', recoverable: true },
          ),
        );
        return;
      }
      throw new OfficeFileViewerError(
        'NAVIGATION_UNSUPPORTED',
        '当前文档暂不支持页面导航',
        { stage: 'navigation', recoverable: false },
      );
    }
    const pageCount = controller.getPageCount();
    const canStillGrow = refs.state.current.document.phase === 'parsing';
    if (
      !Number.isFinite(index) ||
      index < 0 ||
      (!canStillGrow && (!Number.isFinite(pageCount) || index >= pageCount))
    ) {
      clearPendingPageNavigation();
      throw new OfficeFileViewerError(
        'NAVIGATION_OUT_OF_RANGE',
        `页面索引超出范围：${index}`,
        { stage: 'navigation', recoverable: false },
      );
    }
    clearPendingPageNavigation();
    controller.scrollToPage(index);
    void controller.ensurePageMounted(index).catch((error) => {
      const navigationError =
        error instanceof OfficeFileViewerError
          ? error
          : new OfficeFileViewerError(
              'NAVIGATION_OUT_OF_RANGE',
              '目标页面未能及时挂载',
              { stage: 'navigation', recoverable: true, cause: error },
            );
      reportNavigationError(navigationError);
    });
  };

  return {
    openFile: (file) => {
      if (!isHandleActive(refs)) return;
      clearPendingPageNavigation();
      void refs.actions.current.selectFile(file);
    },
    reload: () => {
      if (!isHandleActive(refs)) return;
      clearPendingPageNavigation();
      refs.actions.current.retry();
    },
    setZoom: (value) => {
      if (!hasRenderableDocument(refs)) return;
      refs.actions.current.changeZoom(value);
    },
    zoomIn: () => {
      if (!hasRenderableDocument(refs)) return;
      refs.actions.current.zoomIn();
    },
    zoomOut: () => {
      if (!hasRenderableDocument(refs)) return;
      refs.actions.current.zoomOut();
    },
    fitWidth: () => {
      if (!supportsDocumentAction(refs, 'canFitWidth')) return;
      refs.actions.current.changeZoomMode('fit-width');
    },
    fitPage: () => {
      if (!supportsDocumentAction(refs, 'canFitPage')) return;
      refs.actions.current.changeZoomMode('fit-page');
    },
    goToPage: (index) => {
      if (!isHandleActive(refs)) return;
      if (!supportsDocumentAction(refs, 'canNavigatePages')) return;
      const target = Math.trunc(index);
      if (!Number.isFinite(target) || target < 0) {
        throw new OfficeFileViewerError(
          'NAVIGATION_OUT_OF_RANGE',
          `页面索引超出范围：${index}`,
          { stage: 'navigation', recoverable: false },
        );
      }
      navigatePage(target, true);
    },
    goToSheet: (id) => {
      if (!supportsDocumentAction(refs, 'canNavigateSheets')) return;
      refs.actions.current.selectSheet(id);
    },
    goToSlide: (index) => {
      if (!supportsDocumentAction(refs, 'canNavigateSlides')) return;
      refs.actions.current.selectSlide(index);
    },
    toggleFullscreen: () => {
      if (!supportsDocumentAction(refs, 'canFullscreen')) return;
      void refs.actions.current.toggleFullscreen();
    },
    openSearch: () => {
      if (!supportsDocumentAction(refs, 'canSearch')) return;
      refs.actions.current.openSearch();
    },
    openReview: () => {
      if (!supportsDocumentAction(refs, 'canReview')) return;
      refs.actions.current.openReviewPanel();
    },
    getViewState: () => {
      const view = refs.state.current.view;
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
    },
    getCapabilities: () =>
      isHandleActive(refs) ? getCurrentCapabilities(refs) : undefined,
  };
}
