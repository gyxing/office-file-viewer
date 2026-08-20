// OfficePreviewStage 根据显式舞台状态切换预览组件，并统一处理加载和错误态。
import type { ReactElement } from 'react';
import React, { lazy, memo, Suspense } from 'react';
import { OfficePreviewEmpty } from '../formats/common/OfficePreviewEmpty';
import type { WordRevisionMode } from '../services/annotations/types';
import type { OfficeFileViewerPreviewState } from '../services/parsing/internalTypes';
import type { OfficeFileViewerPresentationMediaOptions } from '../services/presentation/mediaTypes';
import type {
  OfficeFileViewerPresentationTransitions,
  PresentationNavigationIntent,
} from '../services/presentation/transitionTypes';
import type { SpreadsheetViewMode } from '../services/spreadsheet/viewMode';
import { OfficeError } from './Error';
import { OfficeLoading } from './Loading';

const LazyPptxViewer = lazy(() =>
  import('../formats/pptx/PptxViewer').then((module) => ({
    default: module.PptxViewer,
  })),
);
const LazyXlsxViewer = lazy(() =>
  import('../formats/xlsx/XlsxViewer').then((module) => ({
    default: module.XlsxViewer,
  })),
);
const LazyDocxViewer = lazy(() =>
  import('../formats/docx/DocxViewer').then((module) => ({
    default: module.DocxViewer,
  })),
);
const LazyDocViewer = lazy(() =>
  import('../formats/doc/DocViewer').then((module) => ({
    default: module.DocViewer,
  })),
);

/** 演示文稿 Viewer 可以消费的物化或按需预览。 */
type PresentationPreview = Extract<
  OfficeFileViewerPreviewState,
  { previewKind: 'ppt' | 'pptx' }
>;

/** 电子表格 Viewer 可以消费的物化或按需预览。 */
type SpreadsheetPreview = Extract<
  OfficeFileViewerPreviewState,
  { previewKind: 'xls' | 'xlsx' }
>;

/** DOCX Viewer 可以消费的物化或按需预览。 */
type DocxPreview = Extract<
  OfficeFileViewerPreviewState,
  { previewKind: 'docx' }
>;

/** DOC/WPS Viewer 可以消费的物化或按需预览。 */
type DocPreview = Extract<OfficeFileViewerPreviewState, { previewKind: 'doc' }>;

/** 预览舞台所有互斥的可渲染分支。 */
export type OfficePreviewStageState =
  | { kind: 'empty' }
  | {
      kind: 'loading';
      /** 当前解析阶段对应的提示。 */
      tip?: string;
    }
  | {
      kind: 'error';
      /** 阻止继续预览的错误说明。 */
      message: string;
      /** 最近文件来源仍可用时触发重新加载。 */
      retry?: () => void;
    }
  | {
      kind: 'presentation';
      /** 当前演示文稿预览。 */
      preview: PresentationPreview;
      /** 当前选中的幻灯片索引。 */
      activeIndex: number;
      /** 当前预览缩放比例。 */
      zoom: number;
      /** 演讲者备注面板是否展开。 */
      showSpeakerNotes: boolean;
      /** 演示文稿媒体读取配置。 */
      mediaOptions?: false | OfficeFileViewerPresentationMediaOptions;
      /** 是否按源文件播放页级切换。 */
      transitions: OfficeFileViewerPresentationTransitions;
      /** 最近一次工具栏翻页产生的切换意图。 */
      transitionIntent?: PresentationNavigationIntent;
    }
  | {
      kind: 'spreadsheet';
      /** 当前电子表格预览。 */
      preview: SpreadsheetPreview;
      /** 当前选中的工作表标识。 */
      activeSheetId?: string;
      /** 当前预览缩放比例。 */
      zoom: number;
      /** 当前电子表格采用的显示模式。 */
      viewMode: SpreadsheetViewMode;
    }
  | {
      kind: 'docx';
      /** 当前 DOCX 预览。 */
      preview: DocxPreview;
      /** 当前预览缩放比例。 */
      zoom: number;
      /** 文档大纲是否展开。 */
      showOutline: boolean;
      /** 当前采用的 Word 修订内容投影模式。 */
      wordRevisionMode: WordRevisionMode;
    }
  | {
      kind: 'doc';
      /** 当前 DOC/WPS 预览。 */
      preview: DocPreview;
      /** 当前预览缩放比例。 */
      zoom: number;
      /** 文档大纲是否展开。 */
      showOutline: boolean;
      /** 当前采用的 Word 修订内容投影模式。 */
      wordRevisionMode: WordRevisionMode;
    };

/** 预览舞台属性。 */
type OfficePreviewStageProps = {
  /** 当前互斥舞台状态。 */
  state: OfficePreviewStageState;
  /** 关闭文字文档大纲。 */
  onCloseWordOutline: () => void;
  /** 搜索启用时打开文字文档查找侧栏。 */
  onOpenSearch?: () => void;
  /** 选择指定幻灯片。 */
  onSelectSlide: (index: number) => void;
  /** 选择指定工作表。 */
  onSelectSheet: (sheetId: string) => void;
};

/** 根据显式分支选择具体预览器，并保留各格式的懒加载边界。 */
function OfficePreviewStageComponent({
  state,
  onCloseWordOutline,
  onOpenSearch,
  onSelectSlide,
  onSelectSheet,
}: OfficePreviewStageProps) {
  if (state.kind === 'empty') return <OfficePreviewEmpty />;
  if (state.kind === 'loading') return <OfficeLoading tip={state.tip} />;
  if (state.kind === 'error') {
    return <OfficeError message={state.message} onRetry={state.retry} />;
  }

  let content: ReactElement;
  switch (state.kind) {
    case 'presentation':
      content = (
        <LazyPptxViewer
          key={state.preview.sessionId}
          preview={state.preview}
          activeIndex={state.activeIndex}
          zoom={state.zoom}
          showSpeakerNotes={state.showSpeakerNotes}
          mediaOptions={state.mediaOptions}
          transitions={state.transitions}
          transitionIntent={state.transitionIntent}
          onSelectSlide={onSelectSlide}
        />
      );
      break;
    case 'spreadsheet':
      content = (
        <LazyXlsxViewer
          key={state.preview.sessionId}
          preview={state.preview}
          activeSheetId={state.activeSheetId}
          zoom={state.zoom}
          viewMode={state.viewMode}
          onSelectSheet={onSelectSheet}
        />
      );
      break;
    case 'docx':
      content = (
        <LazyDocxViewer
          key={state.preview.sessionId}
          preview={state.preview}
          zoom={state.zoom}
          showOutline={state.showOutline}
          wordRevisionMode={state.wordRevisionMode}
          onCloseOutline={onCloseWordOutline}
          onOpenSearch={onOpenSearch}
        />
      );
      break;
    case 'doc':
      content = (
        <LazyDocViewer
          key={state.preview.sessionId}
          preview={state.preview}
          zoom={state.zoom}
          showOutline={state.showOutline}
          wordRevisionMode={state.wordRevisionMode}
          onCloseOutline={onCloseWordOutline}
          onOpenSearch={onOpenSearch}
        />
      );
      break;
  }

  return <Suspense fallback={<OfficeLoading />}>{content}</Suspense>;
}

export const OfficePreviewStage = memo(OfficePreviewStageComponent);
