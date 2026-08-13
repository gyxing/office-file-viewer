import type { ParsedOfficeFile } from '../../preview';
import {
  disposeDocumentSession,
  type OfficeDocumentSession,
} from '../../session';
import type { PreviewKind } from '../formatDefinitions';
import type {
  MaterializedPreviewState,
  OfficeFileViewerPreviewHandle,
  OfficeFileViewerPreviewState,
} from '../internalTypes';
import type { RuntimeSink } from '../runtime/types';
import type { ParseProgress } from '../types';

/** 格式族会话适配器创建时使用的共享上下文。 */
export type OfficeSessionAdapterContext = {
  /** 当前解析的浏览器文件。 */
  file: File;
  /** 当前文件对应的具体预览格式。 */
  kind: PreviewKind;
  /** 当前会话是否允许发送渐进预览。 */
  enablePartial: boolean;
  /** 统一持有取消信号和解析资源的文档会话。 */
  documentSession: OfficeDocumentSession;
  /** 向解析会话订阅者发送加载进度。 */
  emitProgress(progress: ParseProgress): void;
  /** 向预览器发送格式族已经可以展示的快照。 */
  emitPartial(preview: OfficeFileViewerPreviewState): void;
};

/** 单个格式族向解析会话提供的运行时组装能力。 */
export type OfficeFormatSessionAdapter = {
  /** 接收主线程或 Worker 输出的统一解析事件。 */
  readonly sink: RuntimeSink;
  /** 完成组装并把文档会话所有权交给最终预览句柄。 */
  finish(): Promise<OfficeFileViewerPreviewHandle>;
  /** 解析失败时保留已经可展示的渐进结果。 */
  recoverPartial(): Promise<OfficeFileViewerPreviewState | undefined>;
  /** 释放尚未转移给最终模型的格式资源。 */
  dispose(): void | Promise<void>;
};

/** 动态格式族模块对注册表暴露的创建函数。 */
export type OfficeFormatSessionAdapterFactory = (
  context: OfficeSessionAdapterContext,
) => OfficeFormatSessionAdapter;

/** 创建与具体格式严格对应的物化预览状态。 */
export function createMaterializedPreviewState<Parsed extends ParsedOfficeFile>(
  context: OfficeSessionAdapterContext,
  parsed: Parsed,
): Extract<MaterializedPreviewState, { previewKind: Parsed['kind'] }> {
  return {
    sessionId: context.documentSession.id,
    previewKind: parsed.kind,
    mode: 'materialized',
    model: parsed,
  } as Extract<MaterializedPreviewState, { previewKind: Parsed['kind'] }>;
}

/** 返回物化结果真正持有文档资源的模型对象。 */
function getParsedResultOwner(parsed: ParsedOfficeFile) {
  return parsed.kind === 'xls' || parsed.kind === 'xlsx'
    ? parsed.workbook
    : parsed.document;
}

/** 转移部分结果的资源所有权，并返回不带释放方法的预览快照。 */
export function retainMaterializedPreview(
  context: OfficeSessionAdapterContext,
  parsed: ParsedOfficeFile,
): OfficeFileViewerPreviewState {
  context.documentSession.transferTo(getParsedResultOwner(parsed));
  return createMaterializedPreviewState(context, parsed);
}

/** 转移完整结果的资源所有权，并创建可幂等释放的最终句柄。 */
export function createMaterializedPreviewHandle(
  context: OfficeSessionAdapterContext,
  parsed: ParsedOfficeFile,
): OfficeFileViewerPreviewHandle {
  const owner = getParsedResultOwner(parsed);
  context.documentSession.transferTo(owner);
  return {
    ...createMaterializedPreviewState(context, parsed),
    dispose: () => disposeDocumentSession(owner),
  };
}

type RuntimeSinkOverrides = Partial<Omit<RuntimeSink, 'progress' | 'error'>>;

/** 为格式族未声明的事件生成明确的协议边界错误。 */
function rejectUnexpectedEvent(familyLabel: string, eventLabel: string) {
  return Promise.reject(
    new Error(`${familyLabel}会话收到了不支持的${eventLabel}`),
  );
}

/** 使用统一默认行为补齐格式族不需要处理的运行时事件。 */
export function createOfficeFormatRuntimeSink(
  familyLabel: string,
  context: OfficeSessionAdapterContext,
  overrides: RuntimeSinkOverrides,
): RuntimeSink {
  return {
    progress: context.emitProgress,
    resource:
      overrides.resource ??
      (() => rejectUnexpectedEvent(familyLabel, '资源分块')),
    sheet:
      overrides.sheet ??
      (() => rejectUnexpectedEvent(familyLabel, '工作表分块')),
    spreadsheetMetadata:
      overrides.spreadsheetMetadata ??
      (() => rejectUnexpectedEvent(familyLabel, '工作簿元数据')),
    presentationMetadata:
      overrides.presentationMetadata ??
      (() => rejectUnexpectedEvent(familyLabel, '演示文稿元数据')),
    slide:
      overrides.slide ??
      (() => rejectUnexpectedEvent(familyLabel, '幻灯片分块')),
    documentMetadata:
      overrides.documentMetadata ??
      (() => rejectUnexpectedEvent(familyLabel, 'DOC 元数据')),
    documentBlocks:
      overrides.documentBlocks ??
      (() => rejectUnexpectedEvent(familyLabel, 'DOC 正文分块')),
    docxMetadata:
      overrides.docxMetadata ??
      (() => rejectUnexpectedEvent(familyLabel, 'DOCX 元数据')),
    docxBlocks:
      overrides.docxBlocks ??
      (() => rejectUnexpectedEvent(familyLabel, 'DOCX 正文分块')),
    docxPages:
      overrides.docxPages ??
      (() => rejectUnexpectedEvent(familyLabel, 'DOCX 页面分块')),
    parsed:
      overrides.parsed ??
      (() => rejectUnexpectedEvent(familyLabel, '完整解析结果')),
    complete:
      overrides.complete ??
      (() => rejectUnexpectedEvent(familyLabel, '完成事件')),
    error: () => undefined,
  };
}
