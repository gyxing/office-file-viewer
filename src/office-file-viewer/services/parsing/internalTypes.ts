import type {
  DocWordPageSource,
  DocWordPreviewSummary,
} from '../doc/DocWordPageSource';
import type {
  DocxPagePreviewSource,
  DocxWordPreviewSummary,
} from '../docx/DocxWordPageSource';
import type {
  PresentationSource,
  PresentationSourceSnapshot,
} from '../presentation/PresentationSource';
import type { ParsedOfficeFile } from '../preview';
import type {
  SpreadsheetSource,
  SpreadsheetSourceSnapshot,
} from '../spreadsheet/SpreadsheetSource';
import type { OfficeParseSession } from './types';

/** 按格式判别字段关联物化模型，避免预览类别与模型类型失配。 */
export type MaterializedPreviewState = {
  [Kind in ParsedOfficeFile['kind']]: {
    /** 当前解析或预览会话的标识。 */
    sessionId: string;
    /** 当前文件使用的预览格式类别。 */
    previewKind: Kind;
    /** 当前数据源或渲染器采用的工作模式。 */
    mode: 'materialized';
    /** 当前格式对应的完整解析模型。 */
    model: Extract<ParsedOfficeFile, { kind: Kind }>;
    /** 物化预览不使用按需加载数据源。 */
    source?: undefined;
    /** 物化预览不需要额外摘要。 */
    summary?: undefined;
  };
}[ParsedOfficeFile['kind']];

/** OfficeFileViewer 内部可订阅的非拥有型预览快照。 */
export type OfficeFileViewerPreviewState =
  | MaterializedPreviewState
  | {
      /** 当前解析或预览会话的标识。 */
      sessionId: string;
      /** 当前文件使用的预览格式类别。 */
      previewKind: 'doc';
      /** 当前数据源或渲染器采用的工作模式。 */
      mode: 'source';
      /** 当前预览使用的按需加载数据源。 */
      source: DocWordPageSource;
      /** 当前预览内容的摘要信息。 */
      summary: DocWordPreviewSummary;
      /** 模型。 */
      model?: undefined;
    }
  | {
      /** 当前解析或预览会话的标识。 */
      sessionId: string;
      /** 当前文件使用的预览格式类别。 */
      previewKind: 'docx';
      /** 当前数据源或渲染器采用的工作模式。 */
      mode: 'source';
      /** 当前预览使用的按需加载数据源。 */
      source: DocxPagePreviewSource;
      /** 当前预览内容的摘要信息。 */
      summary: DocxWordPreviewSummary;
      /** 模型。 */
      model?: undefined;
    }
  | {
      /** 当前解析或预览会话的标识。 */
      sessionId: string;
      /** 当前文件使用的预览格式类别。 */
      previewKind: 'ppt' | 'pptx';
      /** 当前数据源或渲染器采用的工作模式。 */
      mode: 'source';
      /** 当前预览使用的按需加载数据源。 */
      source: PresentationSource;
      /** 当前预览内容的摘要信息。 */
      summary: PresentationSourceSnapshot;
      /** 模型。 */
      model?: undefined;
    }
  | {
      /** 当前解析或预览会话的标识。 */
      sessionId: string;
      /** 当前文件使用的预览格式类别。 */
      previewKind: 'xls' | 'xlsx';
      /** 当前数据源或渲染器采用的工作模式。 */
      mode: 'source';
      /** 当前预览使用的按需加载数据源。 */
      source: SpreadsheetSource;
      /** 当前预览内容的摘要信息。 */
      summary: SpreadsheetSourceSnapshot;
      /** 模型。 */
      model?: undefined;
    };

/** OfficeFileViewer 内部最终持有的预览句柄。 */
export type OfficeFileViewerPreviewHandle = OfficeFileViewerPreviewState & {
  /** 幂等释放当前对象持有的资源和订阅。 */
  dispose(): Promise<void>;
};

/** OfficeFileViewer 内部会话支持部分预览，公开解析 API 不暴露该能力。 */
export type OfficeFileViewerParseSession = Omit<
  OfficeParseSession<OfficeFileViewerPreviewHandle>,
  'result'
> & {
  /** 解析会话最终交付的预览句柄。 */
  readonly result: Promise<OfficeFileViewerPreviewHandle>;
  /** 解析尚未完成时可立即展示的最新预览快照。 */
  readonly partialResult: OfficeFileViewerPreviewState | undefined;
  /** 订阅解析过程中的增量结果。 */
  subscribePartial(
    listener: (preview: OfficeFileViewerPreviewState) => void,
  ): () => void;
};
