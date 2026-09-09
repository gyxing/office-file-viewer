import type { DocWordPageSource } from '../services/doc/DocWordPageSource';
import type { DocDocument } from '../services/doc/types';
import type { DocxPagePreviewSource } from '../services/docx/DocxWordPageSource';
import type { DocxDocument } from '../services/docx/types';
import type {
  PreviewFamily,
  PreviewKind,
} from '../services/parsing/formatDefinitions';
import type { PptxDocument } from '../services/pptx/types';
import type { PresentationSource } from '../services/presentation/PresentationSource';
import type { PresentationDocument } from '../services/presentation/types';
import type { SpreadsheetSource } from '../services/spreadsheet/SpreadsheetSource';
import type { SpreadsheetWorkbook } from '../services/spreadsheet/types';

/** 内置格式标识与外部插件可扩展格式标识的联合。 */
export type OfficeFormatId = PreviewKind | (string & Record<never, never>);

/** 描述当前格式可用的预览、导航和导出能力。 */
export type OfficeCapabilities = {
  /** 当前文件对应的预览格式。 */
  previewKind: OfficeFormatId;
  /** 当前格式所属的预览界面族。 */
  family: PreviewFamily;
  /** 是否支持文本查找。 */
  canSearch: boolean;
  /** 是否支持审阅或批注视图。 */
  canReview: boolean;
  /** 是否支持 Word 页面导航。 */
  canNavigatePages: boolean;
  /** 是否支持工作表导航。 */
  canNavigateSheets: boolean;
  /** 是否支持幻灯片导航。 */
  canNavigateSlides: boolean;
  /** 是否支持演讲者备注。 */
  canShowSpeakerNotes: boolean;
  /** 是否支持适合页面缩放。 */
  canFitPage: boolean;
  /** 是否支持适合宽度缩放。 */
  canFitWidth: boolean;
  /** 是否支持浏览器全屏。 */
  canFullscreen: boolean;
  /** 是否可以导出原始文件。 */
  canExportOriginal: boolean;
  /** 当前可用的导出格式。 */
  exportFormats: readonly string[];
  /** 新格式可以追加声明的能力键。 */
  features: Readonly<Record<string, boolean>>;
};

/** 可跨 Worker 传输的资源引用，不携带 Blob 或加载函数。 */
export type OfficeResourceRef = {
  /** 资源在当前文档中的稳定标识。 */
  id: string;
  /** 资源用途类别。 */
  kind: 'image' | 'media' | 'font' | 'other';
  /** 可序列化的资源定位信息。 */
  locator?: OfficeResourceLocator;
  /** 资源 MIME 类型。 */
  mimeType?: string;
  /** 资源声明大小。 */
  size?: number;
};

/** 资源的稳定定位信息，渲染层不会把临时 blob URL 写入快照。 */
export type OfficeResourceLocator =
  | Readonly<{ kind: 'url'; url: string }>
  | Readonly<{ kind: 'lazy'; id: string; mimeType: string; size: number }>;

/** 将源文件对象映射到文档语义节点的位置。 */
export type OfficeSourceMap = Readonly<Record<string, OfficeSourceLocation>>;

/** 描述文档节点在源文件中的位置。 */
export type OfficeSourceLocation = {
  /** OOXML 部件或二进制结构名称。 */
  part?: string;
  /** 源对象标识。 */
  objectId?: string;
  /** 源文件中的偏移量。 */
  offset?: number;
};

/** 文档的通用元数据。 */
export type OfficeDocumentMetadata = {
  /** 原始文件名。 */
  fileName?: string;
  /** 原始文件大小。 */
  fileSize?: number;
  /** 文档标题。 */
  title?: string;
  /** 文档作者。 */
  author?: string;
};

/** 可结构化序列化的扩展数据值。 */
export type OfficeJsonValue =
  | string
  | number
  | boolean
  | null
  | readonly OfficeJsonValue[]
  | Readonly<{ [key: string]: OfficeJsonValue }>;

/** 文档语义节点的稳定标识。 */
export type OfficeNodeId = string;

/** 可跨 Worker 传输的只读文档快照。 */
export type OfficeDocumentSnapshot = Readonly<{
  /** 当前会话内稳定的文档标识。 */
  id: string;
  /** 对外稳定的格式字段。 */
  format: OfficeFormatId;
  /** 当前格式所属的预览界面族。 */
  family: PreviewFamily;
  /** 文档元数据。 */
  metadata: OfficeDocumentMetadata;
  /** 当前文档实际可用的能力。 */
  capabilities: OfficeCapabilities;
  /** 文档使用的资源引用。 */
  resources: readonly OfficeResourceRef[];
  /** 解析过程中产生的可恢复警告。 */
  warnings?: readonly OfficeWarning[];
  /** 语义节点到源文件位置的映射。 */
  sourceMap?: OfficeSourceMap;
  /** 格式插件提供的可序列化扩展数据。 */
  extensionData?: Readonly<Record<string, OfficeJsonValue>>;
}>;

/** 会话内部把快照与物化模型或按需 Source 绑定。 */
export type OfficeDocumentRuntime<
  TModel = OfficeDocumentModel,
  TSource = OfficeDocumentSource,
> =
  | Readonly<{
      /** 公开语义快照。 */
      snapshot: OfficeDocumentSnapshot;
      /** 当前运行时持有完整物化模型。 */
      mode: 'materialized';
      /** 格式专属物化模型。 */
      model: TModel;
      /** 物化模式不持有按需 Source。 */
      source?: never;
    }>
  | Readonly<{
      /** 公开语义快照。 */
      snapshot: OfficeDocumentSnapshot;
      /** 当前运行时持有按需 Source。 */
      mode: 'source';
      /** Source 模式不持有完整模型。 */
      model?: never;
      /** 格式专属按需内容源。 */
      source: TSource;
    }>;

/** 当前内置格式的物化模型联合。 */
export type OfficeDocumentModel =
  | DocDocument
  | DocxDocument
  | SpreadsheetWorkbook
  | PresentationDocument
  | PptxDocument;

/** 当前内置格式的按需内容源联合。 */
export type OfficeDocumentSource =
  | DocWordPageSource
  | DocxPagePreviewSource
  | SpreadsheetSource
  | PresentationSource;

/** 解析过程中的可恢复或诊断警告。 */
export type OfficeWarning = {
  /** 稳定警告代码。 */
  code: string;
  /** 面向调用方的警告说明。 */
  message: string;
  /** 警告来源。 */
  source?: string;
  /** 是否可以继续使用当前结果。 */
  recoverable?: boolean;
};

/** 面向未来编辑器的可序列化变更集合。 */
export type OfficeChangeSet = Readonly<{
  /** 变更集合标识。 */
  id: string;
  /** 变更基于的文档标识。 */
  baseDocumentId: string;
  /** 有序变更操作。 */
  operations: readonly OfficeChangeOperation[];
}>;

/** 文本或节点级变更操作。 */
export type OfficeChangeOperation =
  | Readonly<{
      /** 插入文本。 */
      type: 'insert-text';
      /** 目标节点。 */
      nodeId: OfficeNodeId;
      /** 插入偏移。 */
      offset: number;
      /** 插入内容。 */
      text: string;
    }>
  | Readonly<{
      /** 删除文本。 */
      type: 'delete-text';
      /** 目标节点。 */
      nodeId: OfficeNodeId;
      /** 删除起点。 */
      offset: number;
      /** 删除长度。 */
      length: number;
    }>
  | Readonly<{
      /** 替换节点。 */
      type: 'replace-node';
      /** 目标节点。 */
      nodeId: OfficeNodeId;
      /** JSON-like 节点值。 */
      value: OfficeJsonValue;
    }>;
