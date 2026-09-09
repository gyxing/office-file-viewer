import {
  normalizeOfficeFileViewerError,
  OfficeFileViewerError,
} from '../services/errors/OfficeFileViewerError';
import {
  createOfficeFileViewerParseSession,
  resolveOfficePlugin,
} from '../services/parsing/createParseSession';
import type {
  OfficeFileViewerParseSession,
  OfficeFileViewerPreviewHandle,
  OfficeFileViewerPreviewState,
} from '../services/parsing/internalTypes';
import type {
  OfficeParseOptions,
  OfficeParseSession,
  ParseProgress,
} from '../services/parsing/types';
import type { ParsedOfficeFile } from '../services/preview';
import { createOfficeDocumentSession } from '../services/session/OfficeDocumentSession';
import {
  collectOfficeResourceRefs,
  createOfficeDocumentRuntime,
  createOfficeSourceRuntime,
  createOfficeSourceSnapshot,
  getOfficeCapabilities,
  normalizeOfficeWarnings,
} from './model';
import type {
  OfficeParserPlugin,
  OfficeWorkerParser,
} from './parsingContracts';
import { createOfficeResourceSession } from './resources';
import type { OfficeDocumentRuntime } from './types';

type RuntimePlugin = OfficeParserPlugin & {
  workerSupport?: 'none' | 'main-thread' | 'worker';
  workerFactory?: () => Worker;
  parseInWorker?: OfficeWorkerParser;
};

function createPluginAbortError() {
  const error = new Error('插件解析已取消');
  error.name = 'AbortError';
  return error;
}

/** 支持渐进 Source 结果的公开文档解析会话。 */
export type OfficeDocumentParseSession =
  OfficeParseSession<OfficeDocumentRuntime> & {
    /** 当前已可展示的渐进运行时结果。 */
    readonly partialResult: OfficeDocumentRuntime | undefined;
    /** 订阅渐进运行时结果。 */
    subscribePartial(
      listener: (runtime: OfficeDocumentRuntime) => void,
    ): () => void;
  };

function toRuntime(
  file: File,
  preview: OfficeFileViewerPreviewState,
): OfficeDocumentRuntime {
  if (preview.mode === 'materialized') {
    return createOfficeDocumentRuntime({
      sessionId: preview.sessionId,
      file,
      parsed: preview.model as ParsedOfficeFile,
    });
  }

  const capabilities = getOfficeCapabilities(preview.previewKind, {
    canExportOriginal: true,
  });
  const resources = collectOfficeResourceRefs(preview.summary);
  const warnings = normalizeOfficeWarnings(
    (preview.summary as { warnings?: unknown }).warnings,
  );
  const snapshot = createOfficeSourceSnapshot({
    sessionId: preview.sessionId,
    fileName: file.name,
    fileSize: file.size,
    format: preview.previewKind,
    family:
      preview.previewKind === 'doc' || preview.previewKind === 'docx'
        ? 'word'
        : preview.previewKind === 'xls' || preview.previewKind === 'xlsx'
        ? 'spreadsheet'
        : 'presentation',
    capabilities,
    resources,
    warnings,
  });
  return createOfficeSourceRuntime(snapshot, preview.source);
}

function createPluginDocumentParseSession(
  file: File,
  options: OfficeParseOptions,
): OfficeDocumentParseSession {
  const controller = new AbortController();
  const listeners = new Set<(progress: ParseProgress) => void>();
  const partialListeners = new Set<(runtime: OfficeDocumentRuntime) => void>();
  let status: OfficeParseSession<OfficeDocumentRuntime>['status'] = 'starting';
  let partialResult: OfficeDocumentRuntime | undefined;
  let delegate: OfficeDocumentParseSession | undefined;
  let documentSession:
    | ReturnType<typeof createOfficeDocumentSession>
    | undefined;
  let runtimePlugin: RuntimePlugin | undefined;
  let runtimeResult: OfficeDocumentRuntime | undefined;
  let activeWorker: Worker | undefined;
  let resolvingPlugin = true;

  const emitProgress = (progress: ParseProgress) => {
    listeners.forEach((listener) => {
      try {
        listener(progress);
      } catch {
        // 单个进度订阅者异常不能中断插件解析。
      }
    });
  };
  const emitPartial = (runtime: OfficeDocumentRuntime) => {
    partialResult = runtime;
    partialListeners.forEach((listener) => {
      try {
        listener(runtime);
      } catch {
        // 单个渐进订阅者异常不能中断插件解析。
      }
    });
  };

  const runDelegate = async () => {
    const fallbackOptions: OfficeParseOptions = {
      ...options,
      pluginRegistry: undefined,
    };
    delegate = createOfficeDocumentParseSession(file, fallbackOptions);
    const unsubscribeProgress = delegate.subscribe(emitProgress);
    const unsubscribePartial = delegate.subscribePartial(emitPartial);
    if (controller.signal.aborted) delegate.cancel();
    try {
      const fallbackResult = await delegate.result;
      if (controller.signal.aborted) throw createPluginAbortError();
      status = delegate.status;
      return fallbackResult;
    } finally {
      unsubscribeProgress();
      unsubscribePartial();
    }
  };

  const run = async (): Promise<OfficeDocumentRuntime> => {
    status = 'running';
    try {
      const plugin = await resolveOfficePlugin(
        file,
        options,
        controller.signal,
      );
      resolvingPlugin = false;
      if (!plugin) {
        return runDelegate();
      }

      runtimePlugin = plugin as RuntimePlugin;
      documentSession = createOfficeDocumentSession();
      if (controller.signal.aborted) documentSession.abort();
      const pluginDocumentSession = documentSession;
      const resources = createOfficeResourceSession();
      // 外层解析取消时立即中止资源读取，而不是等到会话完全 dispose 才回收。
      const abortResources = () => {
        void resources.dispose();
      };
      pluginDocumentSession.signal.addEventListener('abort', abortResources, {
        once: true,
      });
      pluginDocumentSession.register({
        dispose: () => {
          pluginDocumentSession.signal.removeEventListener(
            'abort',
            abortResources,
          );
          return resources.dispose();
        },
      });
      const { pluginRegistry: _pluginRegistry, ...parserOptions } = options;
      void _pluginRegistry;
      const { workerFactory: _workerFactory, ...workerParserOptions } =
        parserOptions;
      void _workerFactory;
      const workerMode = options.worker ?? 'auto';
      const isBuiltInPlugin = runtimePlugin.id.startsWith('builtin:');
      if (
        !isBuiltInPlugin &&
        workerMode === 'always' &&
        runtimePlugin.workerSupport !== 'worker'
      ) {
        throw new OfficeFileViewerError(
          'PLUGIN_WORKER_UNSUPPORTED',
          '当前插件不支持强制 Worker 解析',
          { stage: 'plugin', recoverable: false },
        );
      }
      const pluginInput = {
        file,
        sessionId: pluginDocumentSession.id,
        signal: pluginDocumentSession.signal,
        options: parserOptions,
        resources,
        documentSession: pluginDocumentSession,
      };
      runtimeResult = await runtimePlugin.createSource?.(pluginInput);
      if (controller.signal.aborted) throw createPluginAbortError();
      if (runtimeResult) {
        // Source 插件已经负责资源和快照绑定，无需再次进入解析器。
      } else if (runtimePlugin.id.startsWith('builtin:')) {
        await documentSession.dispose();
        documentSession = undefined;
        runtimePlugin = undefined;
        return runDelegate();
      } else if (
        workerMode !== 'never' &&
        runtimePlugin.workerSupport === 'worker' &&
        runtimePlugin.workerFactory &&
        runtimePlugin.parseInWorker
      ) {
        const worker = runtimePlugin.workerFactory();
        activeWorker = worker;
        try {
          runtimeResult = await runtimePlugin.parseInWorker(
            {
              file,
              sessionId: pluginDocumentSession.id,
              options: workerParserOptions,
            },
            worker,
          );
        } finally {
          if (activeWorker === worker) activeWorker = undefined;
          worker.terminate();
        }
      } else {
        runtimeResult = await runtimePlugin.parse(pluginInput);
      }
      if (controller.signal.aborted) throw createPluginAbortError();
      if (!runtimeResult) {
        throw new OfficeFileViewerError(
          'INVALID_FILE',
          '插件没有返回有效的文档运行时结果',
          { stage: 'plugin', recoverable: false },
        );
      }
      documentSession.register({
        dispose: () => runtimePlugin?.dispose?.(runtimeResult!),
      });
      partialResult = runtimeResult;
      emitPartial(runtimeResult);
      status = 'completed';
      return runtimeResult;
    } catch (error) {
      if (controller.signal.aborted) {
        status = 'cancelled';
        delegate?.dispose();
        await documentSession?.dispose();
        throw error;
      }
      status = 'failed';
      await documentSession?.dispose();
      throw normalizeOfficeFileViewerError(error, {
        stage: runtimePlugin || resolvingPlugin ? 'plugin' : 'parsing',
        fileName: file.name,
      });
    }
  };

  const result = run();
  return {
    result,
    get status() {
      return status;
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    cancel() {
      if (status === 'completed' || status === 'cancelled') return;
      controller.abort();
      delegate?.cancel();
      documentSession?.abort();
      activeWorker?.terminate();
    },
    dispose() {
      controller.abort();
      delegate?.dispose();
      void documentSession?.dispose();
      activeWorker?.terminate();
      listeners.clear();
      partialListeners.clear();
    },
    get partialResult() {
      return partialResult;
    },
    subscribePartial(listener) {
      partialListeners.add(listener);
      return () => partialListeners.delete(listener);
    },
  };
}

/** 将现有 Viewer 解析会话转换为可复用的 Core 运行时会话。 */
export function createOfficeDocumentParseSession(
  file: File,
  options: OfficeParseOptions = {},
): OfficeDocumentParseSession {
  if (options.pluginRegistry) {
    return createPluginDocumentParseSession(file, options);
  }
  const internalSession: OfficeFileViewerParseSession =
    createOfficeFileViewerParseSession(file, options);
  let latestHandle: OfficeFileViewerPreviewHandle | undefined;
  let partialResult: OfficeDocumentRuntime | undefined;
  const partialListeners = new Set<(runtime: OfficeDocumentRuntime) => void>();

  const emitPartial = (preview: OfficeFileViewerPreviewState) => {
    partialResult = toRuntime(file, preview);
    partialListeners.forEach((listener) => {
      try {
        listener(partialResult!);
      } catch {
        // 单个渐进订阅者异常不能中断文件解析。
      }
    });
  };

  const unsubscribePartial = internalSession.subscribePartial(emitPartial);
  const result = internalSession.result.then((handle) => {
    latestHandle = handle;
    const runtime = toRuntime(file, handle);
    partialResult = runtime;
    return runtime;
  });

  return {
    result,
    get status() {
      return internalSession.status;
    },
    subscribe(listener: (progress: ParseProgress) => void) {
      return internalSession.subscribe(listener);
    },
    cancel() {
      internalSession.cancel();
    },
    dispose() {
      unsubscribePartial();
      partialListeners.clear();
      internalSession.dispose();
      void latestHandle?.dispose();
    },
    get partialResult() {
      return partialResult;
    },
    subscribePartial(listener) {
      partialListeners.add(listener);
      return () => partialListeners.delete(listener);
    },
  };
}
