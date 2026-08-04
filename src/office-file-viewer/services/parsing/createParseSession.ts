import { DocWordPageSource } from '../doc/DocWordPageSource';
import { OFFICE_LARGE_FILE_THRESHOLDS } from '../performance/officePerformanceThresholds';
import type { ParsedOfficeFile } from '../preview';
import {
  createOfficeDocumentSession,
  disposeDocumentSession,
  type OfficeDocumentSession,
} from '../session';
import {
  DocDocumentAssembler,
  PptDocumentAssembler,
  XlsDocumentAssembler,
} from './assembly/DocumentAssembler';
import { ResourceRegistry } from './assembly/ResourceRegistry';
import { detectPreviewKind } from './detectPreviewKind';
import { tryDetectPreviewKind, type PreviewKind } from './formatDefinitions';
import { loadOfficeSourcePreviewFactory } from './formatParserRegistry';
import type {
  MaterializedPreviewState,
  OfficeFileViewerParseSession,
  OfficeFileViewerPreviewHandle,
  OfficeFileViewerPreviewState,
} from './internalTypes';
import { createRuntime } from './runtime/createRuntime';
import { MainThreadRuntime } from './runtime/MainThreadRuntime';
import type { RuntimeSink } from './runtime/types';
import { isWorkerStartupError, WorkerRuntime } from './runtime/WorkerRuntime';
import type {
  OfficeParseOptions,
  OfficeParseSession,
  OfficeParseSessionStatus,
  ParseProgress,
} from './types';

function createParseSession(
  file: File,
  kind: PreviewKind,
  options: OfficeParseOptions,
  enablePartial: boolean,
  documentSession: OfficeDocumentSession,
): OfficeFileViewerParseSession {
  const listeners = new Set<(progress: ParseProgress) => void>();
  const partialListeners = new Set<
    (preview: OfficeFileViewerPreviewState) => void
  >();
  const useDocPageSource =
    enablePartial &&
    kind === 'doc' &&
    file.size >= OFFICE_LARGE_FILE_THRESHOLDS.cfbFileBytes;
  const assembler =
    kind === 'xls'
      ? new XlsDocumentAssembler(new ResourceRegistry())
      : undefined;
  const presentationAssembler =
    kind === 'ppt'
      ? new PptDocumentAssembler(new ResourceRegistry())
      : undefined;
  const documentAssembler =
    kind === 'doc' && !useDocPageSource
      ? new DocDocumentAssembler(new ResourceRegistry())
      : undefined;
  const docPageSource = useDocPageSource
    ? new DocWordPageSource({
        sessionId: documentSession.id,
        signal: documentSession.signal,
      })
    : undefined;
  let runtime: MainThreadRuntime | WorkerRuntime | undefined;
  let status: OfficeParseSessionStatus = 'starting';
  let parsedResult: ParsedOfficeFile | undefined;
  let partialResult: OfficeFileViewerPreviewState | undefined;
  let ownershipTransferred = false;

  const createMaterializedState = <Parsed extends ParsedOfficeFile>(
    parsed: Parsed,
  ): Extract<MaterializedPreviewState, { previewKind: Parsed['kind'] }> =>
    ({
      sessionId: documentSession.id,
      previewKind: parsed.kind,
      mode: 'materialized',
      model: parsed,
    } as Extract<MaterializedPreviewState, { previewKind: Parsed['kind'] }>);

  const createDocSourceState = (): OfficeFileViewerPreviewState => {
    if (!docPageSource) throw new Error('DOC PageSource 尚未创建');
    return {
      sessionId: documentSession.id,
      previewKind: 'doc',
      mode: 'source',
      source: docPageSource,
      summary: docPageSource.getSummary(),
    };
  };

  const transferParsedResult = (parsed: ParsedOfficeFile) => {
    const owner =
      parsed.kind === 'xls' || parsed.kind === 'xlsx'
        ? parsed.workbook
        : parsed.document;
    documentSession.transferTo(owner);
    ownershipTransferred = true;
    return owner;
  };

  [assembler, presentationAssembler, documentAssembler, docPageSource].forEach(
    (documentAssemblerResource) => {
      if (!documentAssemblerResource) return;
      documentSession.register({
        dispose: () => documentAssemblerResource.dispose(),
      });
    },
  );

  const emitProgress = (progress: ParseProgress) => {
    listeners.forEach((listener) => {
      try {
        listener(progress);
      } catch (listenerError) {
        // 调用方进度回调不能破坏解析任务本身。
        void listenerError;
      }
    });
  };

  const emitPartial = (parsed: ParsedOfficeFile) => {
    if (!enablePartial) return;
    const preview = createMaterializedState(parsed);
    partialListeners.forEach((listener) => {
      try {
        listener(preview);
      } catch (listenerError) {
        // 调用方的渐进渲染异常不能中断底层文件解析。
        void listenerError;
      }
    });
  };

  const emitDocSourcePartial = () => {
    if (!enablePartial || !docPageSource?.hasRenderableContent()) {
      return;
    }
    const preview = createDocSourceState();
    partialListeners.forEach((listener) => {
      try {
        listener(preview);
      } catch (listenerError) {
        // Source 的渐进渲染回调同样不能中断底层解析。
        void listenerError;
      }
    });
  };

  const emitSourcePreviewPartial = (preview: OfficeFileViewerPreviewState) => {
    if (!enablePartial || preview.mode !== 'source') return;
    partialResult = preview;
    partialListeners.forEach((listener) => {
      try {
        listener(preview);
      } catch (listenerError) {
        // Source 的订阅者异常不能中断结构读取、页面测量或按需加载。
        void listenerError;
      }
    });
  };

  const sink: RuntimeSink = {
    progress: emitProgress,
    resource: async (resource) => {
      const target =
        assembler ??
        presentationAssembler ??
        documentAssembler ??
        docPageSource;
      if (!target) throw new Error('当前格式会话收到了资源分块');
      await target.addResource(resource);
    },
    sheet: async (index, revision, sheet) => {
      if (!assembler) throw new Error('非 XLS 会话收到了工作表分块');
      assembler.addSheet(index, revision, sheet);
      if (enablePartial && assembler.hasRenderableContent()) {
        emitPartial({ kind: 'xls', workbook: assembler.snapshot() });
      }
    },
    presentationMetadata: async (metadata) => {
      if (!presentationAssembler) {
        throw new Error('非 PPT 会话收到了演示文稿元数据');
      }
      presentationAssembler.setMetadata(metadata);
      if (presentationAssembler.hasRenderableContent()) {
        emitPartial({
          kind: 'ppt',
          document: presentationAssembler.snapshot(),
        });
      }
    },
    slide: async (index, slide) => {
      if (!presentationAssembler) {
        throw new Error('非 PPT 会话收到了幻灯片分块');
      }
      presentationAssembler.addSlide(index, slide);
      if (presentationAssembler.hasRenderableContent()) {
        emitPartial({
          kind: 'ppt',
          document: presentationAssembler.snapshot(),
        });
      }
    },
    documentMetadata: async (metadata) => {
      if (docPageSource) {
        docPageSource.setMetadata(metadata);
        emitDocSourcePartial();
        return;
      }
      if (!documentAssembler) {
        throw new Error('非 DOC 会话收到了文档元数据');
      }
      documentAssembler.setMetadata(metadata);
      if (documentAssembler.hasRenderableContent()) {
        emitPartial({
          kind: 'doc',
          document: documentAssembler.snapshot(),
        });
      }
    },
    documentBlocks: async (startIndex, blocks) => {
      if (docPageSource) {
        await docPageSource.addBlocks(startIndex, blocks);
        emitDocSourcePartial();
        return;
      }
      if (!documentAssembler) {
        throw new Error('非 DOC 会话收到了正文分块');
      }
      documentAssembler.addBlocks(startIndex, blocks);
      if (documentAssembler.hasRenderableContent()) {
        emitPartial({
          kind: 'doc',
          document: documentAssembler.snapshot(),
        });
      }
    },
    parsed: async (parsed) => {
      parsedResult = parsed;
    },
    complete: async (warnings) => {
      if (assembler) {
        assembler.setWarnings(warnings);
        parsedResult = { kind: 'xls', workbook: assembler.complete() };
      } else if (presentationAssembler) {
        parsedResult = {
          kind: 'ppt',
          document: presentationAssembler.complete(),
        };
      } else if (documentAssembler) {
        parsedResult = {
          kind: 'doc',
          document: documentAssembler.complete(),
        };
      } else if (docPageSource) {
        await docPageSource.complete();
      }
    },
    error: () => undefined,
  };

  const run = async () => {
    status = 'running';
    const workerMode = options.worker ?? 'auto';
    try {
      if (enablePartial) {
        const sourceFactory = await loadOfficeSourcePreviewFactory(kind);
        const sourceHandle = await sourceFactory?.(file, {
          documentSession,
          emitProgress,
          emitPartial: emitSourcePreviewPartial,
        });
        if (sourceHandle) {
          partialResult = sourceHandle;
          ownershipTransferred = true;
          status = 'completed';
          return sourceHandle;
        }
      }

      runtime = createRuntime(workerMode, kind, options.workerFactory);
      if (runtime instanceof WorkerRuntime) {
        try {
          if (kind !== 'xls' && kind !== 'ppt' && kind !== 'doc') {
            throw new Error('当前格式尚未启用 Worker');
          }
          await runtime.run(
            file,
            kind,
            {
              documentSessionId: documentSession.id,
              signal: documentSession.signal,
            },
            sink,
          );
        } catch (error) {
          if (workerMode !== 'auto' || !isWorkerStartupError(error))
            throw error;
          runtime.dispose();
          runtime = new MainThreadRuntime();
          await runtime.run(
            file,
            kind,
            {
              documentSessionId: documentSession.id,
              signal: documentSession.signal,
            },
            sink,
          );
        }
      } else {
        await runtime.run(
          file,
          kind,
          {
            documentSessionId: documentSession.id,
            signal: documentSession.signal,
          },
          sink,
        );
      }
      if (docPageSource) {
        if (!docPageSource.hasRenderableContent()) {
          throw new Error('DOC PageSource 未生成可渲染页面');
        }
        status = 'completed';
        documentSession.transferTo(docPageSource);
        ownershipTransferred = true;
        const state = createDocSourceState();
        const handle: OfficeFileViewerPreviewHandle = {
          ...state,
          dispose: () => disposeDocumentSession(docPageSource),
        };
        return handle;
      }
      if (!parsedResult) throw new Error('解析运行时未返回文档结果');
      status = 'completed';
      const owner = transferParsedResult(parsedResult);
      const state = createMaterializedState(parsedResult);
      const handle: OfficeFileViewerPreviewHandle = {
        ...state,
        dispose: () => disposeDocumentSession(owner),
      };
      return handle;
    } catch (error) {
      const cancelled =
        documentSession.signal.aborted ||
        (error instanceof Error && error.name === 'AbortError');
      if (cancelled) {
        status = 'cancelled';
      } else {
        status = 'failed';
        if (enablePartial && assembler?.hasRenderableContent()) {
          const parsed: ParsedOfficeFile = {
            kind: 'xls',
            workbook: assembler.completePartial(),
          };
          partialResult = createMaterializedState(parsed);
          transferParsedResult(parsed);
        } else if (
          enablePartial &&
          presentationAssembler?.hasRenderableContent()
        ) {
          const parsed: ParsedOfficeFile = {
            kind: 'ppt',
            document: presentationAssembler.completePartial(),
          };
          partialResult = createMaterializedState(parsed);
          transferParsedResult(parsed);
        } else if (enablePartial && documentAssembler?.hasRenderableContent()) {
          const parsed: ParsedOfficeFile = {
            kind: 'doc',
            document: documentAssembler.completePartial(),
          };
          partialResult = createMaterializedState(parsed);
          transferParsedResult(parsed);
        } else if (enablePartial && docPageSource?.hasRenderableContent()) {
          await docPageSource.complete();
          partialResult = createDocSourceState();
          documentSession.transferTo(docPageSource);
          ownershipTransferred = true;
        } else if (enablePartial && partialResult?.mode === 'source') {
          // Source 工厂在首次增量输出前已经转移会话，失败时保留可展示快照。
          ownershipTransferred = true;
        }
      }
      if (!ownershipTransferred) await documentSession.dispose();
      throw error;
    } finally {
      runtime?.dispose();
      runtime = undefined;
    }
  };

  const result = run();
  return {
    result,
    get status() {
      return status;
    },
    get partialResult() {
      return partialResult;
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    subscribePartial(listener) {
      partialListeners.add(listener);
      return () => partialListeners.delete(listener);
    },
    cancel() {
      if (
        status === 'completed' ||
        status === 'cancelled' ||
        status === 'failed'
      ) {
        return;
      }
      documentSession.abort();
      runtime?.dispose();
    },
    dispose() {
      if (status === 'starting' || status === 'running') {
        documentSession.abort();
        runtime?.dispose();
        void documentSession.dispose();
      }
      listeners.clear();
      partialListeners.clear();
    },
  };
}

/** 创建单文件解析会话，统一管理运行时、进度、取消和结果资源。 */
export function createOfficeParseSession(
  file: File,
  options: OfficeParseOptions = {},
): OfficeParseSession<ParsedOfficeFile> {
  const internalSession = createParseSession(
    file,
    detectPreviewKind(file.name),
    options,
    false,
    createOfficeDocumentSession(),
  );
  return {
    result: internalSession.result.then((preview) => {
      if (preview.mode !== 'materialized') {
        throw new Error('公共解析 API 不能返回按需预览源');
      }
      return preview.model;
    }),
    get status() {
      return internalSession.status;
    },
    subscribe: (listener) => internalSession.subscribe(listener),
    cancel: () => internalSession.cancel(),
    dispose: () => internalSession.dispose(),
  };
}

/** 创建仅供 OfficeFileViewer 使用的渐进解析会话。 */
export function createOfficeFileViewerParseSession(
  file: File,
  options: OfficeParseOptions = {},
  documentSession?: OfficeDocumentSession,
): OfficeFileViewerParseSession {
  const kind = tryDetectPreviewKind(file.name);
  if (!kind) throw new Error('暂不支持该文件格式');
  return createParseSession(
    file,
    kind,
    options,
    true,
    documentSession ?? createOfficeDocumentSession(),
  );
}
