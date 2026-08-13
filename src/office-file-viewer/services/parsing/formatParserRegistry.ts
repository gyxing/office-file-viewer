import type { OfficeParseResourcePolicy } from '../../shared/resource/OfficeResourcePolicy';
import type { OfficeDocumentSession } from '../session';
import {
  OFFICE_FORMAT_METADATA,
  type OfficeFormatMetadata,
  type PreviewFamily,
  type PreviewKind,
} from './formatDefinitions';
import type {
  OfficeFileViewerPreviewHandle,
  OfficeFileViewerPreviewState,
} from './internalTypes';
import type { WorkerSourceClient } from './runtime/source/WorkerSourceClient';
import type { RuntimeContext, RuntimeSink } from './runtime/types';
import type { OfficeFormatSessionAdapterFactory } from './sessionAdapters/types';
import type { ParseProgress } from './types';

/** 格式解析器向统一运行时输出增量或完整结果的入口合同。 */
export type OfficeFormatParser = (
  file: File,
  context: RuntimeContext,
  sink: RuntimeSink,
) => Promise<void>;

/** 提前创建按需预览源时可使用的会话与事件接口。 */
export type OfficeSourcePreviewContext = {
  /** 持有取消信号并统一管理 Reader、Source 等资源的文档会话。 */
  documentSession: OfficeDocumentSession;
  /** 向解析会话发送当前格式的加载进度。 */
  emitProgress(progress: ParseProgress): void;
  /** 向界面发送已经可以展示的按需预览快照。 */
  emitPartial(preview: OfficeFileViewerPreviewState): void;
  /** 当前解析会话采用的可选资源上限。 */
  resourcePolicy?: OfficeParseResourcePolicy;
  /** 大 OOXML 文件需要在 Worker 长期持有 Reader 时使用的可选客户端。 */
  workerSourceClient?: WorkerSourceClient;
};

/** 在大文件满足按需加载条件时创建预览源，否则交回普通解析流程。 */
export type OfficeSourcePreviewFactory = (
  file: File,
  context: OfficeSourcePreviewContext,
) => Promise<OfficeFileViewerPreviewHandle | undefined>;

/** 单种格式的稳定能力及其按需加载入口。 */
export type OfficeFormatDefinition = OfficeFormatMetadata & {
  /** 动态加载当前格式的主线程解析入口。 */
  loadParser: () => Promise<OfficeFormatParser>;
  /** 动态加载当前格式可选的大文件预览源工厂。 */
  loadSourcePreviewFactory?: () => Promise<OfficeSourcePreviewFactory>;
};

/** 六种格式仅在实际选中后加载对应解析模块，避免主入口聚合全部解析代码。 */
const OFFICE_FORMAT_DEFINITIONS = {
  doc: {
    ...OFFICE_FORMAT_METADATA.doc,
    loadParser: () =>
      import('../doc/parseDoc').then(({ runDocParser }) => runDocParser),
  },
  docx: {
    ...OFFICE_FORMAT_METADATA.docx,
    loadParser: () =>
      import('../docx/parseDocx').then(({ runDocxParser }) => runDocxParser),
    loadSourcePreviewFactory: () =>
      import('../docx/parseDocxSource').then(
        ({ tryCreateDocxSourcePreview }) => tryCreateDocxSourcePreview,
      ),
  },
  ppt: {
    ...OFFICE_FORMAT_METADATA.ppt,
    loadParser: () =>
      import('../ppt/parsePpt').then(({ runPptParser }) => runPptParser),
    loadSourcePreviewFactory: () =>
      import('../ppt/PptPresentationSource').then(
        ({ tryCreatePptSourcePreview }) => tryCreatePptSourcePreview,
      ),
  },
  pptx: {
    ...OFFICE_FORMAT_METADATA.pptx,
    loadParser: () =>
      import('../pptx/parsePptx').then(({ runPptxParser }) => runPptxParser),
    loadSourcePreviewFactory: () =>
      import('../pptx/PptxPresentationSource').then(
        ({ tryCreatePptxSourcePreview }) => tryCreatePptxSourcePreview,
      ),
  },
  xls: {
    ...OFFICE_FORMAT_METADATA.xls,
    loadParser: () =>
      import('../xls/parseXls').then(({ runXlsParser }) => runXlsParser),
    loadSourcePreviewFactory: () =>
      import('../xls/XlsSpreadsheetSource').then(
        ({ tryCreateXlsSourcePreview }) => tryCreateXlsSourcePreview,
      ),
  },
  xlsx: {
    ...OFFICE_FORMAT_METADATA.xlsx,
    loadParser: () =>
      import('../xlsx/parseXlsx').then(({ runXlsxParser }) => runXlsxParser),
    loadSourcePreviewFactory: () =>
      import('../xlsx/XlsxSpreadsheetSource').then(
        ({ tryCreateXlsxSourcePreview }) => tryCreateXlsxSourcePreview,
      ),
  },
} satisfies Record<PreviewKind, OfficeFormatDefinition>;

/** 三类预览族只在普通解析链路真正启动时加载对应会话组装逻辑。 */
const OFFICE_SESSION_ADAPTER_LOADERS = {
  word: () =>
    import('./sessionAdapters/createWordSessionAdapter').then(
      ({ createWordSessionAdapter }) => createWordSessionAdapter,
    ),
  spreadsheet: () =>
    import('./sessionAdapters/createSpreadsheetSessionAdapter').then(
      ({ createSpreadsheetSessionAdapter }) => createSpreadsheetSessionAdapter,
    ),
  presentation: () =>
    import('./sessionAdapters/createPresentationSessionAdapter').then(
      ({ createPresentationSessionAdapter }) =>
        createPresentationSessionAdapter,
    ),
} satisfies Record<
  PreviewFamily,
  () => Promise<OfficeFormatSessionAdapterFactory>
>;

/** 返回目标格式的完整动态能力定义。 */
export function getOfficeFormatDefinition(
  kind: PreviewKind,
): OfficeFormatDefinition {
  return OFFICE_FORMAT_DEFINITIONS[kind];
}

/** 只加载当前文件格式对应的主线程解析入口。 */
export async function loadOfficeFormatParser(
  kind: PreviewKind,
): Promise<OfficeFormatParser> {
  return getOfficeFormatDefinition(kind).loadParser();
}

/** 按需加载当前格式的大文件预览源工厂。 */
export async function loadOfficeSourcePreviewFactory(
  kind: PreviewKind,
): Promise<OfficeSourcePreviewFactory | undefined> {
  return getOfficeFormatDefinition(kind).loadSourcePreviewFactory?.();
}

/** 按具体格式所属的预览族加载内部会话组装适配器。 */
export async function loadOfficeSessionAdapterFactory(
  kind: PreviewKind,
): Promise<OfficeFormatSessionAdapterFactory> {
  const family = getOfficeFormatDefinition(kind).family;
  return OFFICE_SESSION_ADAPTER_LOADERS[family]();
}
