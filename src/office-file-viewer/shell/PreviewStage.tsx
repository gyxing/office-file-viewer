// OfficePreviewStage 根据当前文件格式切换到对应预览组件，并统一处理加载和错误态。
import React, { lazy, memo, Suspense } from 'react';
import type {
  DocWordPageSource,
  DocWordPreviewSummary,
} from '../services/doc/DocWordPageSource';
import type { DocDocument } from '../services/doc/types';
import type {
  DocxWordPageSource,
  DocxWordPreviewSummary,
} from '../services/docx/DocxWordPageSource';
import type { DocxDocument } from '../services/docx/types';
import type {
  PresentationDocument,
  PresentationSource,
} from '../services/presentation/types';
import {
  isSpreadsheetPreviewKind,
  type PreviewKind,
} from '../services/preview';
import type { SpreadsheetSource } from '../services/spreadsheet/SpreadsheetSource';
import type { SpreadsheetWorkbook } from '../services/spreadsheet/types';
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

/** 定义 OfficePreviewStage 组件可接收的属性。 */
type OfficePreviewStageProps = {
  /** 文件当前是否仍在加载或解析。 */
  loading: boolean;
  /** OfficePreviewStageProps 的 loadingTip 文本值。 */
  loadingTip?: string;
  /** 当前是否已有可交付渲染器显示的内容。 */
  hasRenderableContent: boolean;
  /** OfficePreviewStageProps 携带的结构化解析错误。 */
  error?: string;
  /** 当前文件识别出的预览格式。 */
  previewKind: PreviewKind;
  /** 当前文件解析 Session，用于隔离格式内部的渐进状态。 */
  documentSessionId?: string;
  /** 已标准化的 PPTX 演示文稿模型；未提供时使用来源格式或渲染器的默认行为。 */
  pptxDocument?: PresentationDocument;
  /** 大型 PPT/PPTX 使用的按页读取数据源。 */
  presentationPreviewSource?: PresentationSource;
  /** 已标准化的 XLS/XLSX 工作簿模型；未提供时使用来源格式或渲染器的默认行为。 */
  spreadsheetWorkbook?: SpreadsheetWorkbook;
  /** 大型 XLS/XLSX 使用的按 Sheet 数据源。 */
  spreadsheetPreviewSource?: SpreadsheetSource;
  /** 已标准化的 DOCX 文档模型；未提供时使用来源格式或渲染器的默认行为。 */
  docxDocument?: DocxDocument;
  /** 大型 DOCX 使用的流式页面来源。 */
  docxPreviewSource?: DocxWordPageSource;
  /** 大型 DOCX 不含完整 blocks/pages 的轻量摘要。 */
  docxPreviewSummary?: DocxWordPreviewSummary;
  /** 已标准化的 DOC/WPS 文档模型；未提供时使用来源格式或渲染器的默认行为。 */
  docDocument?: DocDocument;
  /** 大型 DOC/WPS 使用的渐进页面来源。 */
  docPreviewSource?: DocWordPageSource;
  /** 大型 DOC/WPS 不含完整正文的轻量摘要。 */
  docPreviewSummary?: DocWordPreviewSummary;
  /** 当前选中项在所属集合中的索引。 */
  activeIndex: number;
  /** OfficePreviewStageProps 的 activeSheetId 文本值。 */
  activeSheetId?: string;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 演讲者备注面板当前是否展开。 */
  showSpeakerNotes: boolean;
  /** 在 SelectSlide 事件发生时调用的回调函数。 */
  onSelectSlide: (index: number) => void;
  /** 在 SelectSheet 事件发生时调用的回调函数。 */
  onSelectSheet: (sheetId: string) => void;
};

/** 渲染 OfficePreviewStageComponent 组件。 */
function OfficePreviewStageComponent({
  loading,
  loadingTip,
  hasRenderableContent,
  error,
  previewKind,
  documentSessionId,
  pptxDocument,
  presentationPreviewSource,
  spreadsheetWorkbook,
  spreadsheetPreviewSource,
  docxDocument,
  docxPreviewSource,
  docxPreviewSummary,
  docDocument,
  docPreviewSource,
  docPreviewSummary,
  activeIndex,
  activeSheetId,
  zoom,
  showSpeakerNotes,
  onSelectSlide,
  onSelectSheet,
}: OfficePreviewStageProps) {
  if (error) return <OfficeError message={error} />;
  if (loading && !hasRenderableContent) {
    return <OfficeLoading tip={loadingTip} />;
  }

  // 格式 viewer 是真正的重渲染模块，按文件类型懒加载，避免首屏一次性拉取所有预览实现。
  return (
    <Suspense fallback={<OfficeLoading />}>
      {isSpreadsheetPreviewKind(previewKind) ? (
        <LazyXlsxViewer
          workbook={spreadsheetWorkbook}
          source={spreadsheetPreviewSource}
          kind={previewKind}
          activeSheetId={activeSheetId}
          zoom={zoom}
          onSelectSheet={onSelectSheet}
        />
      ) : previewKind === 'docx' ? (
        <LazyDocxViewer
          document={docxDocument}
          source={docxPreviewSource}
          summary={docxPreviewSummary}
          zoom={zoom}
          documentSessionId={documentSessionId ?? 'word-unloaded'}
        />
      ) : previewKind === 'doc' ? (
        <LazyDocViewer
          document={docDocument}
          source={docPreviewSource}
          summary={docPreviewSummary}
          zoom={zoom}
          documentSessionId={documentSessionId ?? 'word-unloaded'}
        />
      ) : (
        <LazyPptxViewer
          document={pptxDocument}
          source={presentationPreviewSource}
          activeIndex={activeIndex}
          zoom={zoom}
          showSpeakerNotes={showSpeakerNotes}
          onSelectSlide={onSelectSlide}
        />
      )}
    </Suspense>
  );
}

export const OfficePreviewStage = memo(OfficePreviewStageComponent);
