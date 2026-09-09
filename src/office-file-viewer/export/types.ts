import type {
  OfficeChangeSet,
  OfficeDocumentSnapshot,
  OfficeWarning,
} from '../core/types';
import type { OfficeFileViewerError } from '../services/errors/OfficeFileViewerError';

/**
 * 请求把只读快照或变更集合导出为指定格式。
 * @experimental 导出协议首版。
 */
export type OfficeExportRequest = {
  /** 待导出的文档快照。 */
  document: OfficeDocumentSnapshot;
  /** 可选的编辑变更集合。 */
  changes?: OfficeChangeSet;
  /** 目标格式，缺省时使用快照格式。 */
  format?: string;
  /** 取消导出和资源读取。 */
  signal?: AbortSignal;
  /** 是否尽量保留未知源文件部件。 */
  preserveUnknownParts?: boolean;
};

/** 导出成功后的二进制结果。 */
export type OfficeExportResult = {
  /** 导出的文件 Blob。 */
  blob: Blob;
  /** 建议下载文件名。 */
  fileName: string;
  /** 实际输出格式。 */
  format: string;
  /** 导出过程中的可恢复警告。 */
  warnings?: readonly OfficeWarning[];
};

/**
 * 修改后文件的格式写入器。
 * @experimental 导出器协议首版。
 */
export type OfficeExporter = {
  /** 写入器稳定标识。 */
  id: string;
  /** 写入器支持的格式或扩展名。 */
  formats: readonly string[];
  /** 执行一次导出。 */
  export(request: OfficeExportRequest): Promise<OfficeExportResult>;
};

/** 导出阶段可由调用方分支处理的错误代码。 */
export type OfficeExportErrorCode =
  | 'EXPORT_UNSUPPORTED'
  | 'EXPORT_CANCELLED'
  | 'EXPORT_FILENAME_REQUIRED'
  | 'EXPORT_RESOURCE_FAILED'
  | 'EXPORT_WRITE_FAILED';

/** 导出错误沿用 OfficeFileViewer 的结构化错误类型。 */
export type OfficeExportError = OfficeFileViewerError;

/**
 * 导出器的实例级注册表。
 * @experimental 导出注册表首版。
 */
export type OfficeExporterRegistry = {
  /** 返回冻结的导出器列表。 */
  list(): readonly OfficeExporter[];
  /** 按格式或扩展名查询导出器。 */
  getByFormat(format: string): OfficeExporter | undefined;
  /** 注册导出器并返回幂等注销函数。 */
  register(exporter: OfficeExporter): () => void;
  /** 清理注册表索引。 */
  dispose(): void;
};

/** 原始文件导出的可接受输入。 */
export type OfficeOriginalSource = File | Blob | ArrayBuffer | Uint8Array;
