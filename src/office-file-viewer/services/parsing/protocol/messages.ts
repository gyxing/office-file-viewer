import type { OfficeParseResourcePolicy } from '../../../shared/resource/OfficeResourcePolicy';
import type { DocBlock, DocDocument } from '../../doc/types';
import type {
  DocxBlock,
  DocxDocument,
  DocxPageContent,
} from '../../docx/types';
import type { PresentationSourceSnapshot } from '../../presentation/PresentationSource';
import type {
  PresentationDocument,
  SlideModel,
} from '../../presentation/types';
import type {
  OfficeSearchProgress,
  OfficeSearchQuery,
} from '../../search/types';
import type {
  SpreadsheetSheetLayout,
  SpreadsheetSourceSnapshot,
} from '../../spreadsheet/SpreadsheetSource';
import type { SpreadsheetPerformanceProfile } from '../../spreadsheet/spreadsheetPerformance';
import type {
  SpreadsheetSheet,
  SpreadsheetWarning,
  SpreadsheetWorkbook,
} from '../../spreadsheet/types';
import type { WordPageSourceSnapshot } from '../../word/WordPageSource';
import type { WordOutlineItem, WordPerformanceProfile } from '../../word/types';
import type { PreviewKind } from '../formatDefinitions';
import type { ParseProgress } from '../types';
import type { SerializedParseError } from './errors';

/** 保存跨线程解析协议主体内容之外的元数据。 */
export type PortablePresentationMetadata = Omit<
  PresentationDocument,
  'slides' | 'resources'
>;

/** 保存跨线程解析协议主体内容之外的元数据。 */
export type PortableDocMetadata = Omit<
  DocDocument,
  'blocks' | 'paragraphs' | 'resources'
>;

/** 保存 DOCX 跨线程正文和页面分块之外的文档元数据。 */
export type PortableDocxMetadata = Omit<DocxDocument, 'blocks' | 'pages'>;

/** 保存电子表格跨线程工作表分块之外的工作簿元数据。 */
export type PortableSpreadsheetMetadata = Omit<
  SpreadsheetWorkbook,
  'sheets' | 'warnings' | 'resources'
>;

/** 可通过结构化克隆跨线程传输的图片资源。 */
export type PortableResource =
  | {
      /** 在所属集合中的唯一标识。 */
      id: string;
      /** 资源在跨线程消息中采用的编码形式。 */
      encoding: 'binary';
      /** 资源的 MIME 类型，用于选择解码和渲染方式。 */
      mimeType: string;
      /** 资源或文件的二进制缓冲区；发送方移交后不再继续使用。 */
      buffer: ArrayBuffer;
    }
  | {
      /** 在所属集合中的唯一标识。 */
      id: string;
      /** 资源在跨线程消息中采用的编码形式。 */
      encoding: 'text';
      /** 资源的 MIME 类型，用于选择解码和渲染方式。 */
      mimeType: 'image/svg+xml';
      /** 文本内容。 */
      text: string;
    }
  | {
      /** 在所属集合中的唯一标识。 */
      id: string;
      /** 资源在跨线程消息中采用的编码形式。 */
      encoding: 'rgba';
      /** 资源的 MIME 类型，用于选择解码和渲染方式。 */
      mimeType: 'image/png';
      /** 宽度，单位为标准化渲染像素。 */
      width: number;
      /** 高度，单位为标准化渲染像素。 */
      height: number;
      /** 资源或文件的二进制缓冲区；发送方移交后不再继续使用。 */
      buffer: ArrayBuffer;
    };

/** 定义跨线程解析协议中传输的消息结构。 */
type ParseMainToWorkerMessage =
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-start';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 当前模型对应的 Office 内容类型。 */
      kind: PreviewKind;
      /** 正在解析的原始文件名，用于格式识别和错误提示。 */
      fileName: string;
      /** 当前任务中由发送方递增的消息序号。 */
      sequence: number;
      /** 交由 Worker 内部读取的原始文件，避免主线程提前物化缓冲区。 */
      file: File;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-cancel';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 当前任务中由发送方递增的消息序号。 */
      sequence: number;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-ack';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 分块消息的递增序号，用于 ACK 背压和顺序校验。 */
      sequence: number;
    };

/** 定义跨线程解析协议中传输的消息结构。 */
type ParseWorkerToMainMessage =
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'worker-ready';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-progress';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 当前任务中由发送方递增的消息序号。 */
      sequence: number;
      /** 当前解析阶段及其完成度信息。 */
      progress: ParseProgress;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-docx-meta';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 分块消息的递增序号，用于 ACK 背压和顺序校验。 */
      sequence: number;
      /** DOCX 正文和页面集合之外的文档元数据。 */
      metadata: PortableDocxMetadata;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-docx-blocks';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 分块消息的递增序号，用于 ACK 背压和顺序校验。 */
      sequence: number;
      /** 分块在完整正文集合中的起始索引。 */
      startIndex: number;
      /** 按源文档顺序排列的 DOCX 正文块。 */
      blocks: DocxBlock[];
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-docx-pages';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 分块消息的递增序号，用于 ACK 背压和顺序校验。 */
      sequence: number;
      /** 分块在完整页面集合中的起始索引。 */
      startIndex: number;
      /** 按源文档顺序排列的 DOCX 页面。 */
      pages: DocxPageContent[];
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-resource';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 分块消息的递增序号，用于 ACK 背压和顺序校验。 */
      sequence: number;
      /** 解析器产生并等待主线程确认接收的可移植资源分块。 */
      resource: PortableResource;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-spreadsheet-meta';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 分块消息的递增序号，用于 ACK 背压和顺序校验。 */
      sequence: number;
      /** 工作表集合之外的工作簿级元数据。 */
      metadata: PortableSpreadsheetMetadata;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-sheet';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 分块消息的递增序号，用于 ACK 背压和顺序校验。 */
      sequence: number;
      /** 工作表在工作簿集合中的索引。 */
      sheetIndex: number;
      /** 数据源变更时递增的修订号。 */
      revision: number;
      /** 当前处理的工作表。 */
      sheet: SpreadsheetSheet;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-presentation-meta';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 分块消息的递增序号，用于 ACK 背压和顺序校验。 */
      sequence: number;
      /** 主体元数据，不包含后续分块传输的大型内容。 */
      metadata: PortablePresentationMetadata;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-slide';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 分块消息的递增序号，用于 ACK 背压和顺序校验。 */
      sequence: number;
      /** 幻灯片在演示文稿集合中的索引。 */
      slideIndex: number;
      /** 当前处理或展示的幻灯片。 */
      slide: SlideModel;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-document-meta';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 分块消息的递增序号，用于 ACK 背压和顺序校验。 */
      sequence: number;
      /** 主体元数据，不包含后续分块传输的大型内容。 */
      metadata: PortableDocMetadata;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-document-blocks';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 分块消息的递增序号，用于 ACK 背压和顺序校验。 */
      sequence: number;
      /** 分块在完整集合中的起始索引。 */
      startIndex: number;
      /** 按源文档顺序排列的内容块。 */
      blocks: DocBlock[];
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-complete';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 当前任务中由发送方递增的消息序号。 */
      sequence: number;
      /** 解析时产生但不阻止继续预览的警告。 */
      warnings?: SpreadsheetWarning[];
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-error';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 当前任务中由发送方递增的消息序号。 */
      sequence: number;
      /** 当前操作产生的错误；未提供表示没有错误。 */
      error: SerializedParseError;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'parse-cancelled';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 解析任务的唯一标识，用于过滤其他 Worker 任务的消息。 */
      taskId: string;
      /** 文档会话标识，用于拒绝旧会话或其他 Viewer 的消息。 */
      documentSessionId: string;
      /** 当前任务中由发送方递增的消息序号。 */
      sequence: number;
    };

/** 仅 OOXML 格式需要长期驻留 Worker 的按需数据源。 */
export type WorkerSourceKind = Extract<PreviewKind, 'docx' | 'xlsx' | 'pptx'>;

/** Worker Word Source 初次就绪和后续更新时传输的轻量状态。 */
export type WorkerWordSourceState = {
  /** 当前代理对应的格式。 */
  kind: 'docx';
  /** 文档摘要；资源字段已替换为可移植标记。 */
  summary: unknown;
  /** 页面描述符快照，不包含正文页面。 */
  snapshot: WordPageSourceSnapshot;
  /** 当前已发现的大纲条目。 */
  outlineItems: WordOutlineItem[];
  /** 当前大纲扫描是否结束。 */
  outlineComplete: boolean;
  /** 当前文档采用的性能配置。 */
  performance: WordPerformanceProfile;
  /** 当前待测量批次；资源字段已替换为可移植标记。 */
  measurementBatch?: unknown;
};

/** Worker Spreadsheet Source 初次就绪和后续更新时传输的轻量状态。 */
export type WorkerSpreadsheetSourceState = {
  /** 当前代理对应的格式。 */
  kind: 'xlsx';
  /** 工作簿和 Sheet 描述符快照。 */
  snapshot: SpreadsheetSourceSnapshot;
  /** 各 Sheet 当前性能配置。 */
  profiles: Record<string, SpreadsheetPerformanceProfile>;
  /** 各 Sheet 当前行列布局。 */
  layouts: Record<string, SpreadsheetSheetLayout>;
};

/** Worker Presentation Source 初次就绪和后续更新时传输的轻量状态。 */
export type WorkerPresentationSourceState = {
  /** 当前代理对应的格式。 */
  kind: 'pptx';
  /** 演示文稿和幻灯片描述符快照。 */
  snapshot: PresentationSourceSnapshot;
};

/** 长期 Worker Source 可推送到主线程的状态联合。 */
export type WorkerSourceState =
  | WorkerWordSourceState
  | WorkerSpreadsheetSourceState
  | WorkerPresentationSourceState;

/** Source 打开结果；画像未命中大文件条件时不创建长期代理。 */
export type WorkerSourceOpenResult =
  | { available: false }
  | { available: true; source: WorkerSourceState };

/** 主线程发往长期 Source Worker 的消息。 */
type SourceMainToWorkerMessage =
  | {
      /** 请求 Worker 打开并持有 OOXML Source。 */
      type: 'source-open';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 当前长期 Worker 任务的唯一标识。 */
      taskId: string;
      /** 文档会话标识，用于过滤切换文件后的旧消息。 */
      documentSessionId: string;
      /** 当前请求的递增标识。 */
      requestId: number;
      /** 需要创建按需数据源的 OOXML 格式。 */
      kind: WorkerSourceKind;
      /** 原始文件名。 */
      fileName: string;
      /** Worker 直接读取的原始文件。 */
      file: File;
      /** 宿主可选的资源保护策略。 */
      resourcePolicy?: OfficeParseResourcePolicy;
    }
  | {
      /** 调用长期 Source 的单个按需操作。 */
      type: 'source-request';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 当前长期 Worker 任务的唯一标识。 */
      taskId: string;
      /** 文档会话标识，用于过滤切换文件后的旧消息。 */
      documentSessionId: string;
      /** 当前请求的递增标识。 */
      requestId: number;
      /** Source 内部需要执行的稳定操作名。 */
      operation: string;
      /** 操作参数；必须可被结构化克隆。 */
      args?: unknown;
    }
  | {
      /** 取消指定的未完成 Source 请求。 */
      type: 'source-request-cancel';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 当前长期 Worker 任务的唯一标识。 */
      taskId: string;
      /** 文档会话标识，用于过滤切换文件后的旧消息。 */
      documentSessionId: string;
      /** 需要取消的请求标识。 */
      requestId: number;
    }
  | {
      /** 释放 Worker 持有的 Source、Reader 和缓存。 */
      type: 'source-dispose';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 当前长期 Worker 任务的唯一标识。 */
      taskId: string;
      /** 文档会话标识，用于过滤切换文件后的旧消息。 */
      documentSessionId: string;
      /** 释放请求的递增标识。 */
      requestId: number;
    };

/** 长期 Source Worker 发往主线程的消息。 */
type SourceWorkerToMainMessage =
  | {
      /** Worker Source 已完成画像和初始化。 */
      type: 'source-opened';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 当前长期 Worker 任务的唯一标识。 */
      taskId: string;
      /** 文档会话标识，用于过滤切换文件后的旧消息。 */
      documentSessionId: string;
      /** 当前响应对应的请求标识。 */
      requestId: number;
      /** Source 初始化结果。 */
      result: WorkerSourceOpenResult;
    }
  | {
      /** 单个 Source RPC 已成功完成。 */
      type: 'source-response';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 当前长期 Worker 任务的唯一标识。 */
      taskId: string;
      /** 文档会话标识，用于过滤切换文件后的旧消息。 */
      documentSessionId: string;
      /** 当前响应对应的请求标识。 */
      requestId: number;
      /** RPC 返回值；资源字段已替换为可移植标记。 */
      result?: unknown;
    }
  | {
      /** Source RPC 或初始化失败。 */
      type: 'source-error';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 当前长期 Worker 任务的唯一标识。 */
      taskId: string;
      /** 文档会话标识，用于过滤切换文件后的旧消息。 */
      documentSessionId: string;
      /** 当前错误对应的请求标识。 */
      requestId: number;
      /** 可跨线程传输的结构化错误。 */
      error: SerializedParseError;
    }
  | {
      /** Source 的轻量快照或待测量批次发生变化。 */
      type: 'source-update';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 当前长期 Worker 任务的唯一标识。 */
      taskId: string;
      /** 文档会话标识，用于过滤切换文件后的旧消息。 */
      documentSessionId: string;
      /** 当前 Source 最新的轻量状态。 */
      source: WorkerSourceState;
    }
  | {
      /** Worker 正在执行 Source 初始化或按需操作。 */
      type: 'source-progress';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 当前长期 Worker 任务的唯一标识。 */
      taskId: string;
      /** 文档会话标识，用于过滤切换文件后的旧消息。 */
      documentSessionId: string;
      /** 当前进度所属请求；初始化阶段为 open 请求。 */
      requestId: number;
      /** 当前解析阶段及其完成度信息。 */
      progress: ParseProgress;
    }
  | {
      /** 大文件搜索在 Worker 中产生一个增量批次。 */
      type: 'source-search-progress';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 当前长期 Worker 任务的唯一标识。 */
      taskId: string;
      /** 文档会话标识，用于过滤切换文件后的旧消息。 */
      documentSessionId: string;
      /** 当前搜索 RPC 的请求标识。 */
      requestId: number;
      /** 搜索查询，用于调试和隔离旧结果。 */
      query: OfficeSearchQuery;
      /** 本批搜索进度和全部新增结果。 */
      progress: OfficeSearchProgress;
    }
  | {
      /** Worker 后台 Source 解析在打开后发生不可恢复错误。 */
      type: 'source-failed';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 当前长期 Worker 任务的唯一标识。 */
      taskId: string;
      /** 文档会话标识，用于过滤切换文件后的旧消息。 */
      documentSessionId: string;
      /** 可跨线程传输的结构化错误。 */
      error: SerializedParseError;
    }
  | {
      /** Worker 已释放长期 Source。 */
      type: 'source-disposed';
      /** 消息或数据结构采用的协议版本号。 */
      version: number;
      /** 当前长期 Worker 任务的唯一标识。 */
      taskId: string;
      /** 文档会话标识，用于过滤切换文件后的旧消息。 */
      documentSessionId: string;
      /** 当前响应对应的释放请求标识。 */
      requestId: number;
    };

/** 主线程能够发送给解析 Worker 的全部消息。 */
export type MainToWorkerMessage =
  | ParseMainToWorkerMessage
  | SourceMainToWorkerMessage;

/** 解析 Worker 能够发送给主线程的全部消息。 */
export type WorkerToMainMessage =
  | ParseWorkerToMainMessage
  | SourceWorkerToMainMessage;
