import type { OfficeParseResourcePolicy } from '../../../../shared/resource/OfficeResourcePolicy';
import type { OfficeResourceSource } from '../../../resource-store/types';
import type {
  OfficeSearchProgress,
  OfficeSearchQuery,
} from '../../../search/types';
import { deserializeParseError } from '../../protocol/errors';
import type {
  MainToWorkerMessage,
  WorkerSourceKind,
  WorkerSourceOpenResult,
  WorkerSourceState,
  WorkerToMainMessage,
} from '../../protocol/messages';
import { OFFICE_PARSER_PROTOCOL_VERSION } from '../../protocol/version';
import type { ParseProgress } from '../../types';

let sourceTaskSequence = 0;

/** Worker Source 入口加载和协议握手允许等待的默认时长。 */
const DEFAULT_SOURCE_READY_TIMEOUT_MS = 10_000;
/** dispose 等待 Worker 主动释放 Reader 和缓存的最长时长。 */
const SOURCE_DISPOSE_GRACE_MS = 2_000;

/** Worker 中懒资源跨线程传输时使用的可移植标记。 */
type WorkerLazyResourceMarker = {
  /** 保持资源模型的判别字段，便于递归识别。 */
  kind: 'lazy';
  /** 当前模型在文档中的稳定资源标识。 */
  id: string;
  /** 资源负载的 MIME 类型。 */
  mimeType: string;
  /** 当前数据占用的空间大小。 */
  size: number;
  /** Worker 资源注册表中的稳定标识。 */
  workerResourceId: string;
};

/** 单个 RPC 的完成函数和可选增量回调。 */
type PendingRequest = {
  /** 完成当前 RPC。 */
  resolve(value: unknown): void;
  /** 拒绝当前 RPC。 */
  reject(error: unknown): void;
  /** 取消监听清理函数。 */
  cleanupAbort(): void;
  /** 接收 Source 初始化或按需读取进度。 */
  onProgress?(progress: ParseProgress): void;
  /** 接收大文件搜索的增量结果。 */
  onSearchProgress?(progress: OfficeSearchProgress): void;
};

/** 单次 Worker Source RPC 的可选控制项。 */
export type WorkerSourceRequestOptions = {
  /** 取消调用方等待，并通知 Worker 取消对应请求。 */
  signal?: AbortSignal;
  /** 接收初始化或按需读取进度。 */
  onProgress?(progress: ParseProgress): void;
  /** 接收 Worker 搜索分批返回的结果。 */
  onSearchProgress?(progress: OfficeSearchProgress): void;
};

/** 创建带稳定错误码的长期 Worker Source 错误。 */
function createWorkerSourceError(code: string, message: string) {
  const error = new Error(message) as Error & {
    /** 供调用方识别当前失败类别的稳定代码。 */
    code: string;
    /** 是否允许 auto 模式回退到主线程。 */
    recoverable: boolean;
  };
  error.name = 'OfficeWorkerSourceError';
  error.code = code;
  error.recoverable = code === 'WORKER_STARTUP_FAILED';
  return error;
}

function isWorkerLazyResourceMarker(
  value: unknown,
): value is WorkerLazyResourceMarker {
  if (!value || typeof value !== 'object') return false;
  const candidate = value as Partial<WorkerLazyResourceMarker>;
  return (
    candidate.kind === 'lazy' &&
    typeof candidate.workerResourceId === 'string' &&
    typeof candidate.id === 'string' &&
    typeof candidate.mimeType === 'string' &&
    typeof candidate.size === 'number'
  );
}

/** 管理长期 Source Worker 的握手、RPC、增量事件和资源生命周期。 */
export class WorkerSourceClient {
  private readonly taskId = `office-source-${Date.now()}-${++sourceTaskSequence}`;
  private readonly pending = new Map<number, PendingRequest>();
  private readonly updateListeners = new Set<
    (source: WorkerSourceState) => void
  >();
  private readonly failureListeners = new Set<(error: Error) => void>();
  private readonly resourceCache = new Map<string, OfficeResourceSource>();
  private worker: Worker | undefined;
  private readyPromise: Promise<void> | undefined;
  private resolveReady: (() => void) | undefined;
  private rejectReady: ((error: unknown) => void) | undefined;
  private readyTimeoutId: ReturnType<typeof setTimeout> | undefined;
  private requestSequence = 0;
  private disposing = false;
  private disposed = false;
  private disposePromise?: Promise<void>;

  constructor(
    readonly documentSessionId: string,
    private readonly workerFactory?: () => Worker,
    private readonly readyTimeoutMs = DEFAULT_SOURCE_READY_TIMEOUT_MS,
  ) {}

  /** 打开并持有目标 OOXML Source；画像未命中时返回不可用结果。 */
  async openSource(
    file: File,
    kind: WorkerSourceKind,
    resourcePolicy?: OfficeParseResourcePolicy,
    options: WorkerSourceRequestOptions = {},
  ) {
    await this.ensureReady();
    this.throwIfUnavailable();
    const requestId = this.nextRequestId();
    const result = await this.sendRequest<WorkerSourceOpenResult>(
      {
        type: 'source-open',
        version: OFFICE_PARSER_PROTOCOL_VERSION,
        taskId: this.taskId,
        documentSessionId: this.documentSessionId,
        requestId,
        kind,
        fileName: file.name,
        file,
        resourcePolicy,
      },
      options,
    );
    if (!result.available) return result;
    return {
      available: true as const,
      source: this.hydrateValue(result.source) as WorkerSourceState,
    };
  }

  /** 调用 Worker 内 Source 的单个按需操作。 */
  async request<TResult = unknown>(
    operation: string,
    args?: unknown,
    options: WorkerSourceRequestOptions = {},
  ): Promise<TResult> {
    await this.ensureReady();
    this.throwIfUnavailable();
    const requestId = this.nextRequestId();
    const result = await this.sendRequest<TResult>(
      {
        type: 'source-request',
        version: OFFICE_PARSER_PROTOCOL_VERSION,
        taskId: this.taskId,
        documentSessionId: this.documentSessionId,
        requestId,
        operation,
        args,
      },
      options,
    );
    return this.hydrateValue(result) as TResult;
  }

  /** 在 Worker 内执行搜索，正文结果仅按批次返回主线程。 */
  search(
    query: OfficeSearchQuery,
    emit: (progress: OfficeSearchProgress) => void,
    signal: AbortSignal,
  ) {
    return this.request<void>(
      'search',
      { query },
      {
        signal,
        onSearchProgress: emit,
      },
    );
  }

  /** 在 Worker 内保留可视范围，并返回对应的异步释放函数。 */
  retain(operation: string, args: unknown) {
    const retentionId = `${this.taskId}:retain:${this.nextRequestId()}`;
    void this.request(operation, { ...this.asRecord(args), retentionId }).catch(
      () => undefined,
    );
    let released = false;
    return () => {
      if (released) return;
      released = true;
      void this.request('release-retention', { retentionId }).catch(
        () => undefined,
      );
    };
  }

  /** 订阅 Worker Source 的轻量快照变化。 */
  subscribe(listener: (source: WorkerSourceState) => void) {
    this.updateListeners.add(listener);
    return () => this.updateListeners.delete(listener);
  }

  /** 订阅 Source 打开后发生的后台不可恢复错误。 */
  subscribeFailure(listener: (error: Error) => void) {
    this.failureListeners.add(listener);
    return () => this.failureListeners.delete(listener);
  }

  /** 幂等释放 Worker、Reader、缓存、订阅和所有未完成 RPC。 */
  dispose() {
    if (this.disposePromise) return this.disposePromise;
    this.disposing = true;
    const disposedError = createWorkerSourceError(
      'WORKER_SOURCE_DISPOSED',
      'Worker Source 已释放',
    );
    this.rejectPending(disposedError);
    this.disposePromise = (async () => {
      const worker = this.worker;
      if (worker) {
        const requestId = this.nextRequestId();
        const acknowledged = new Promise<void>((resolve) => {
          this.pending.set(requestId, {
            resolve: () => resolve(),
            reject: () => resolve(),
            cleanupAbort: () => undefined,
          });
        });
        const message: MainToWorkerMessage = {
          type: 'source-dispose',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId: this.taskId,
          documentSessionId: this.documentSessionId,
          requestId,
        };
        worker.postMessage(message);
        await Promise.race([
          acknowledged,
          new Promise<void>((resolve) => {
            setTimeout(resolve, SOURCE_DISPOSE_GRACE_MS);
          }),
        ]);
      }
      this.cleanupWorker();
      this.rejectPending(disposedError);
      this.resourceCache.clear();
      this.updateListeners.clear();
      this.failureListeners.clear();
      this.disposed = true;
    })();
    return this.disposePromise;
  }

  private nextRequestId() {
    this.requestSequence += 1;
    return this.requestSequence;
  }

  private throwIfUnavailable() {
    if (!this.disposed && !this.disposing) return;
    throw createWorkerSourceError(
      'WORKER_SOURCE_DISPOSED',
      'Worker Source 已释放',
    );
  }

  private ensureReady() {
    if (this.disposed || this.disposing) {
      return Promise.reject(
        createWorkerSourceError(
          'WORKER_SOURCE_DISPOSED',
          'Worker Source 已释放',
        ),
      );
    }
    if (this.readyPromise) return this.readyPromise;
    this.readyPromise = new Promise<void>((resolve, reject) => {
      this.resolveReady = resolve;
      this.rejectReady = reject;
    });
    try {
      this.worker = this.workerFactory
        ? this.workerFactory()
        : new Worker(new URL('../worker/entry.js', import.meta.url), {
            type: 'module',
            name: 'office-file-viewer-source',
          });
    } catch {
      const error = createWorkerSourceError(
        'WORKER_STARTUP_FAILED',
        '无法创建 Office Source Worker',
      );
      this.rejectReady?.(error);
      this.cleanupWorker();
      return this.readyPromise;
    }
    this.worker.addEventListener('message', this.handleMessage);
    this.worker.addEventListener('error', this.handleWorkerError);
    this.readyTimeoutId = setTimeout(() => {
      const error = createWorkerSourceError(
        'WORKER_STARTUP_FAILED',
        'Office Source Worker 就绪超时',
      );
      this.rejectReady?.(error);
      this.rejectPending(error);
      this.cleanupWorker();
    }, Math.max(0, this.readyTimeoutMs));
    return this.readyPromise;
  }

  private sendRequest<TResult>(
    message: MainToWorkerMessage,
    options: WorkerSourceRequestOptions,
  ) {
    if (this.disposed || this.disposing) {
      return Promise.reject(
        createWorkerSourceError(
          'WORKER_SOURCE_DISPOSED',
          'Worker Source 已释放',
        ),
      );
    }
    if (!('requestId' in message)) {
      return Promise.reject(new Error('Source RPC 缺少 requestId'));
    }
    if (options.signal?.aborted) {
      return Promise.reject(this.createAbortError());
    }
    return new Promise<TResult>((resolve, reject) => {
      const abort = () => {
        this.pending.delete(message.requestId);
        const cancelMessage: MainToWorkerMessage = {
          type: 'source-request-cancel',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId: this.taskId,
          documentSessionId: this.documentSessionId,
          requestId: message.requestId,
        };
        this.worker?.postMessage(cancelMessage);
        reject(this.createAbortError());
      };
      const cleanupAbort = () =>
        options.signal?.removeEventListener('abort', abort);
      options.signal?.addEventListener('abort', abort, { once: true });
      this.pending.set(message.requestId, {
        resolve: (value) => resolve(value as TResult),
        reject,
        cleanupAbort,
        onProgress: options.onProgress,
        onSearchProgress: options.onSearchProgress,
      });
      this.worker?.postMessage(message);
    });
  }

  private readonly handleMessage = (
    event: MessageEvent<WorkerToMainMessage>,
  ) => {
    const message = event.data;
    if (message.version !== OFFICE_PARSER_PROTOCOL_VERSION) return;
    if (message.type === 'worker-ready') {
      if (this.readyTimeoutId !== undefined) clearTimeout(this.readyTimeoutId);
      this.resolveReady?.();
      this.resolveReady = undefined;
      this.rejectReady = undefined;
      return;
    }
    if (
      message.taskId !== this.taskId ||
      message.documentSessionId !== this.documentSessionId
    ) {
      return;
    }
    if (message.type === 'source-update') {
      const source = this.hydrateValue(message.source) as WorkerSourceState;
      this.updateListeners.forEach((listener) => listener(source));
      return;
    }
    if (message.type === 'source-failed') {
      const error = deserializeParseError(message.error);
      this.failureListeners.forEach((listener) => listener(error));
      this.rejectPending(error);
      return;
    }
    if (message.type === 'source-progress') {
      this.pending.get(message.requestId)?.onProgress?.(message.progress);
      return;
    }
    if (message.type === 'source-search-progress') {
      this.pending.get(message.requestId)?.onSearchProgress?.(message.progress);
      return;
    }
    if (
      message.type === 'source-opened' ||
      message.type === 'source-response' ||
      message.type === 'source-disposed'
    ) {
      const pending = this.takePending(message.requestId);
      pending?.resolve('result' in message ? message.result : undefined);
      return;
    }
    if (message.type === 'source-error') {
      this.takePending(message.requestId)?.reject(
        deserializeParseError(message.error),
      );
    }
  };

  private readonly handleWorkerError = () => {
    const error = createWorkerSourceError(
      this.resolveReady ? 'WORKER_STARTUP_FAILED' : 'WORKER_RUNTIME_CRASH',
      this.resolveReady
        ? 'Office Source Worker 加载失败'
        : 'Office Source Worker 运行异常',
    );
    this.rejectReady?.(error);
    this.rejectPending(error);
    this.failureListeners.forEach((listener) => listener(error));
    this.cleanupWorker();
  };

  private takePending(requestId: number) {
    const pending = this.pending.get(requestId);
    this.pending.delete(requestId);
    pending?.cleanupAbort();
    return pending;
  }

  private rejectPending(error: unknown) {
    this.pending.forEach((pending) => {
      pending.cleanupAbort();
      pending.reject(error);
    });
    this.pending.clear();
  }

  private cleanupWorker() {
    if (this.readyTimeoutId !== undefined) clearTimeout(this.readyTimeoutId);
    this.readyTimeoutId = undefined;
    this.worker?.removeEventListener('message', this.handleMessage);
    this.worker?.removeEventListener('error', this.handleWorkerError);
    this.worker?.terminate();
    this.worker = undefined;
  }

  private hydrateValue(value: unknown): unknown {
    if (isWorkerLazyResourceMarker(value)) {
      const cached = this.resourceCache.get(value.workerResourceId);
      if (cached) return cached;
      const source: OfficeResourceSource = {
        kind: 'lazy',
        id: value.id,
        mimeType: value.mimeType,
        size: value.size,
        load: async (signal) => {
          const result = await this.request<{
            /** 资源实际采用的 MIME 类型。 */
            mimeType: string;
            /** Worker 返回的二进制负载。 */
            buffer: ArrayBuffer;
          }>(
            'load-resource',
            { resourceId: value.workerResourceId },
            { signal },
          );
          return new Blob([result.buffer], { type: result.mimeType });
        },
      };
      this.resourceCache.set(value.workerResourceId, source);
      return source;
    }
    if (Array.isArray(value)) {
      return value.map((item) => this.hydrateValue(item));
    }
    if (!value || typeof value !== 'object') return value;
    if (
      value instanceof ArrayBuffer ||
      ArrayBuffer.isView(value) ||
      (typeof Blob !== 'undefined' && value instanceof Blob)
    ) {
      return value;
    }
    return Object.fromEntries(
      Object.entries(value).map(([key, item]) => [
        key,
        this.hydrateValue(item),
      ]),
    );
  }

  private createAbortError() {
    const error = new Error('Worker Source 操作已取消');
    error.name = 'AbortError';
    return error;
  }

  private asRecord(value: unknown): Record<string, unknown> {
    return value && typeof value === 'object'
      ? (value as Record<string, unknown>)
      : {};
  }
}
