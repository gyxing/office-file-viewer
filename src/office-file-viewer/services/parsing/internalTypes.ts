import type {
  DocWordPageSource,
  DocWordPreviewSummary,
} from '../doc/DocWordPageSource';
import type {
  DocxWordPageSource,
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

/** OfficeFileViewer 内部可订阅的非拥有型预览快照。 */
export type OfficeFileViewerPreviewState =
  | {
      /** 当前解析或预览会话的标识。 */
      sessionId: string;
      /** 当前文件使用的预览格式类别。 */
      previewKind: ParsedOfficeFile['kind'];
      /** 当前数据源或渲染器采用的工作模式。 */
      mode: 'materialized';
      /** 模型。 */
      model: ParsedOfficeFile;
      /** 当前预览使用的按需加载数据源。 */
      source?: undefined;
      /** 当前预览内容的摘要信息。 */
      summary?: undefined;
    }
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
      source: DocxWordPageSource;
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
  /** DOC 域计算结果的文本片段。 */
  readonly result: Promise<OfficeFileViewerPreviewHandle>;
  /** 解析尚未完成时可立即展示的最新预览快照。 */
  readonly partialResult: OfficeFileViewerPreviewState | undefined;
  /** 订阅解析过程中的增量结果。 */
  subscribePartial(
    listener: (preview: OfficeFileViewerPreviewState) => void,
  ): () => void;
};
