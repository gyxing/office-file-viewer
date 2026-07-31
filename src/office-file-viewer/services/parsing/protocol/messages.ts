import type { DocBlock, DocDocument } from '../../doc/types';
import type {
  PresentationDocument,
  SlideModel,
} from '../../presentation/types';
import type {
  SpreadsheetSheet,
  SpreadsheetWarning,
} from '../../spreadsheet/types';
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
export type MainToWorkerMessage =
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
      kind: 'xls' | 'ppt' | 'doc';
      /** 正在解析的原始文件名，用于格式识别和错误提示。 */
      fileName: string;
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
export type WorkerToMainMessage =
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
      /** 当前解析阶段及其完成度信息。 */
      progress: ParseProgress;
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
    };
