import type { ParsedOfficeFile } from '../preview';
import type { OfficeParseSession } from './types';

/** OfficeFileViewer 内部会话可订阅非拥有型快照，公开解析 API 不暴露该能力。 */
export type OfficeFileViewerParseSession =
  OfficeParseSession<ParsedOfficeFile> & {
    /** OfficeFileViewerParseSession 关联的 partialResult 结构；字段形状由 ParsedOfficeFile | undefined 定义。 */
    readonly partialResult: ParsedOfficeFile | undefined;
    /** 订阅解析过程中的非拥有型部分结果。 */
    subscribePartial(listener: (parsed: ParsedOfficeFile) => void): () => void;
  };
