import { OfficeResourceLimitError } from '../../shared/resource/OfficeResourceLimitError';
import { validateOfficeResourcePolicy } from '../../shared/resource/OfficeResourcePolicy';
import { OFFICE_LARGE_FILE_THRESHOLDS } from '../performance/officePerformanceThresholds';
import type { ParsedOfficeFile } from '../preview';
import {
  createOfficeDocumentSession,
  type OfficeDocumentSession,
} from '../session';
import { detectPreviewKind } from './detectPreviewKind';
import { tryDetectPreviewKind, type PreviewKind } from './formatDefinitions';
import {
  loadOfficeSessionAdapterFactory,
  loadOfficeSourcePreviewFactory,
} from './formatParserRegistry';
import type {
  OfficeFileViewerParseSession,
  OfficeFileViewerPreviewHandle,
  OfficeFileViewerPreviewState,
} from './internalTypes';
import { createRuntime } from './runtime/createRuntime';
import { MainThreadRuntime } from './runtime/MainThreadRuntime';
import { WorkerSourceClient } from './runtime/source/WorkerSourceClient';
import { isWorkerStartupError, WorkerRuntime } from './runtime/WorkerRuntime';
import type { OfficeFormatSessionAdapter } from './sessionAdapters/types';
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
  let runtime: MainThreadRuntime | WorkerRuntime | undefined;
  let sessionAdapter: OfficeFormatSessionAdapter | undefined;
  let workerSourceClient: WorkerSourceClient | undefined;
  let unregisterWorkerSource: (() => void) | undefined;
  let status: OfficeParseSessionStatus = 'starting';
  let partialResult: OfficeFileViewerPreviewState | undefined;
  let ownershipTransferred = false;

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

  const emitPartial = (preview: OfficeFileViewerPreviewState) => {
    if (!enablePartial) return;
    partialListeners.forEach((listener) => {
      try {
        listener(preview);
      } catch (listenerError) {
        // 调用方的渐进渲染异常不能中断底层文件解析。
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

  const run = async () => {
    status = 'running';
    const workerMode = options.worker ?? 'auto';
    const resourcePolicy = options.resourcePolicy;
    let timeoutId: ReturnType<typeof setTimeout> | undefined;
    try {
      validateOfficeResourcePolicy(resourcePolicy);
      if (
        resourcePolicy?.maxFileBytes !== undefined &&
        file.size > resourcePolicy.maxFileBytes
      ) {
        throw new OfficeResourceLimitError(
          'FILE_SIZE_LIMIT_EXCEEDED',
          'Office 文件大小超过宿主配置的上限',
          { limit: resourcePolicy.maxFileBytes, actual: file.size },
        );
      }
      if (resourcePolicy?.timeoutMs !== undefined) {
        timeoutId = setTimeout(() => {
          documentSession.abort(
            new OfficeResourceLimitError(
              'PARSE_TIMEOUT',
              'Office 文件解析超过宿主配置的时限',
              { limit: resourcePolicy.timeoutMs },
            ),
          );
        }, resourcePolicy.timeoutMs);
      }
      if (enablePartial) {
        const sourceFactory = await loadOfficeSourcePreviewFactory(kind);
        const isOoxml = kind === 'docx' || kind === 'xlsx' || kind === 'pptx';
        if (sourceFactory && isOoxml && workerMode !== 'never') {
          workerSourceClient = new WorkerSourceClient(
            documentSession.id,
            options.workerFactory,
          );
          unregisterWorkerSource = documentSession.register({
            dispose: () => workerSourceClient?.dispose(),
          });
        }
        const sourceContext = {
          documentSession,
          emitProgress,
          emitPartial: emitSourcePreviewPartial,
          resourcePolicy,
          workerSourceClient,
        };
        let sourceHandle: OfficeFileViewerPreviewHandle | undefined;
        try {
          sourceHandle = await sourceFactory?.(file, sourceContext);
        } catch (error) {
          if (
            workerMode !== 'auto' ||
            !workerSourceClient ||
            !isWorkerStartupError(error)
          ) {
            throw error;
          }
          await workerSourceClient.dispose();
          unregisterWorkerSource?.();
          unregisterWorkerSource = undefined;
          workerSourceClient = undefined;
          sourceHandle = await sourceFactory?.(file, {
            documentSession,
            emitProgress,
            emitPartial: emitSourcePreviewPartial,
            resourcePolicy,
          });
        }
        if (sourceHandle) {
          partialResult = sourceHandle;
          ownershipTransferred = true;
          status = 'completed';
          return sourceHandle;
        }
        if (workerSourceClient) {
          await workerSourceClient.dispose();
          unregisterWorkerSource?.();
          unregisterWorkerSource = undefined;
          workerSourceClient = undefined;
        }
      }

      const adapterFactory = await loadOfficeSessionAdapterFactory(kind);
      sessionAdapter = adapterFactory({
        file,
        kind,
        enablePartial,
        documentSession,
        emitProgress,
        emitPartial,
      });
      documentSession.register({
        dispose: () => sessionAdapter?.dispose(),
      });

      const isOoxml = kind === 'docx' || kind === 'xlsx' || kind === 'pptx';
      const runtimeMode =
        workerMode === 'auto' &&
        isOoxml &&
        file.size < OFFICE_LARGE_FILE_THRESHOLDS.ooxmlCompressedBytes
          ? 'never'
          : workerMode;
      runtime = createRuntime(runtimeMode, kind, options.workerFactory);
      if (runtime instanceof WorkerRuntime) {
        try {
          await runtime.run(
            file,
            kind,
            {
              documentSessionId: documentSession.id,
              signal: documentSession.signal,
              resourcePolicy,
            },
            sessionAdapter.sink,
          );
        } catch (error) {
          if (runtimeMode !== 'auto' || !isWorkerStartupError(error))
            throw error;
          runtime.dispose();
          runtime = new MainThreadRuntime();
          await runtime.run(
            file,
            kind,
            {
              documentSessionId: documentSession.id,
              signal: documentSession.signal,
              resourcePolicy,
            },
            sessionAdapter.sink,
          );
        }
      } else {
        await runtime.run(
          file,
          kind,
          {
            documentSessionId: documentSession.id,
            signal: documentSession.signal,
            resourcePolicy,
          },
          sessionAdapter.sink,
        );
      }
      const handle = await sessionAdapter.finish();
      status = 'completed';
      ownershipTransferred = true;
      return handle;
    } catch (error) {
      const cancelled =
        documentSession.signal.aborted ||
        (error instanceof Error && error.name === 'AbortError');
      if (cancelled) {
        status = 'cancelled';
      } else {
        status = 'failed';
        const recoveredPartial = enablePartial
          ? await sessionAdapter?.recoverPartial()
          : undefined;
        if (recoveredPartial) {
          partialResult = recoveredPartial;
          ownershipTransferred = true;
        } else if (enablePartial && partialResult?.mode === 'source') {
          // Source 工厂在首次增量输出前已经转移会话，失败时保留可展示快照。
          ownershipTransferred = true;
        }
      }
      // 清理会触发内部 AbortSignal，必须先保存真实取消原因，避免覆盖原始解析错误。
      const abortReason = documentSession.signal.aborted
        ? documentSession.signal.reason
        : undefined;
      if (!ownershipTransferred) await documentSession.dispose();
      throw abortReason ?? error;
    } finally {
      if (timeoutId !== undefined) clearTimeout(timeoutId);
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
