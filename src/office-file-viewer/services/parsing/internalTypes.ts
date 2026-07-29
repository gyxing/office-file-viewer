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
      sessionId: string;
      previewKind: ParsedOfficeFile['kind'];
      mode: 'materialized';
      model: ParsedOfficeFile;
      source?: undefined;
      summary?: undefined;
    }
  | {
      sessionId: string;
      previewKind: 'doc';
      mode: 'source';
      source: DocWordPageSource;
      summary: DocWordPreviewSummary;
      model?: undefined;
    }
  | {
      sessionId: string;
      previewKind: 'docx';
      mode: 'source';
      source: DocxWordPageSource;
      summary: DocxWordPreviewSummary;
      model?: undefined;
    }
  | {
      sessionId: string;
      previewKind: 'ppt' | 'pptx';
      mode: 'source';
      source: PresentationSource;
      summary: PresentationSourceSnapshot;
      model?: undefined;
    }
  | {
      sessionId: string;
      previewKind: 'xls' | 'xlsx';
      mode: 'source';
      source: SpreadsheetSource;
      summary: SpreadsheetSourceSnapshot;
      model?: undefined;
    };

/** OfficeFileViewer 内部最终持有的预览句柄。 */
export type OfficeFileViewerPreviewHandle = OfficeFileViewerPreviewState & {
  dispose(): Promise<void>;
};

/** OfficeFileViewer 内部会话支持部分预览，公开解析 API 不暴露该能力。 */
export type OfficeFileViewerParseSession = Omit<
  OfficeParseSession<OfficeFileViewerPreviewHandle>,
  'result'
> & {
  readonly result: Promise<OfficeFileViewerPreviewHandle>;
  readonly partialResult: OfficeFileViewerPreviewState | undefined;
  subscribePartial(
    listener: (preview: OfficeFileViewerPreviewState) => void,
  ): () => void;
};
