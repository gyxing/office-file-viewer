import type { PreviewKind } from '../formatDefinitions';
import { deserializeParseError } from '../protocol/errors';
import type {
  MainToWorkerMessage,
  WorkerToMainMessage,
} from '../protocol/messages';
import { OFFICE_PARSER_PROTOCOL_VERSION } from '../protocol/version';
import type { RuntimeContext, RuntimeSink } from './types';
import { createParseAbortError } from './types';

let taskSequence = 0;

/** Worker 入口加载和协议握手允许等待的默认时长。 */
const DEFAULT_WORKER_READY_TIMEOUT_MS = 10_000;

/** 描述解析运行时失败时使用的结构化错误。 */
type RuntimeError = Error & {
  /** 供程序识别当前情况的稳定代码。 */
  code: string;
  /** 表示该错误是否允许自动降级到其他解析方式。 */
  recoverable: boolean;
};

function createRuntimeError(
  code: string,
  message: string,
  recoverable: boolean,
): RuntimeError {
  const error = new Error(message) as RuntimeError;
  error.name = 'OfficeWorkerError';
  error.code = code;
  error.recoverable = recoverable;
  return error;
}

/** 判断错误是否发生在文件缓冲区移交 Worker 之前。 */
export function isWorkerStartupError(error: unknown) {
  return (
    error instanceof Error &&
    'code' in error &&
    (
      error as {
        /** Worker 序列化错误对象的稳定代码，用于程序化识别具体情况。 */
        code?: unknown;
      }
    ).code === 'WORKER_STARTUP_FAILED'
  );
}

/** 管理单次解析使用的独立 Worker 和跨线程消息。 */
export class WorkerRuntime {
  private worker: Worker | undefined;
  private stopActive: (() => void) | undefined;

  constructor(
    private readonly workerFactory?: () => Worker,
    private readonly readyTimeoutMs = DEFAULT_WORKER_READY_TIMEOUT_MS,
  ) {}

  run(
    file: File,
    kind: PreviewKind,
    context: RuntimeContext,
    sink: RuntimeSink,
  ): Promise<void> {
    const { documentSessionId, signal } = context;
    if (this.worker) {
      return Promise.reject(
        createRuntimeError(
          'WORKER_BUSY',
          '解析 Worker 正在处理其他任务',
          false,
        ),
      );
    }

    return new Promise<void>((resolve, reject) => {
      let worker: Worker;
      try {
        worker = this.workerFactory
          ? this.workerFactory()
          : new Worker(new URL('./worker/entry.js', import.meta.url), {
              type: 'module',
              name: 'office-file-viewer-parser',
            });
      } catch {
        reject(
          createRuntimeError(
            'WORKER_STARTUP_FAILED',
            '无法创建 Office 解析 Worker',
            true,
          ),
        );
        return;
      }

      const taskId = `office-parse-${Date.now()}-${++taskSequence}`;
      let parseStarted = false;
      let settled = false;
      let readyTimeoutId: ReturnType<typeof setTimeout> | undefined;
      this.worker = worker;

      const cleanup = () => {
        if (readyTimeoutId !== undefined) clearTimeout(readyTimeoutId);
        signal.removeEventListener('abort', handleAbort);
        worker.removeEventListener('message', handleMessage);
        worker.removeEventListener('error', handleWorkerError);
        worker.terminate();
        if (this.worker === worker) this.worker = undefined;
        this.stopActive = undefined;
      };
      const fail = (error: unknown) => {
        if (settled) return;
        settled = true;
        cleanup();
        reject(error);
      };
      const finish = () => {
        if (settled) return;
        settled = true;
        cleanup();
        resolve();
      };
      const sendAck = (sequence: number) => {
        const message: MainToWorkerMessage = {
          type: 'parse-ack',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId,
          documentSessionId,
          sequence,
        };
        worker.postMessage(message);
      };
      function handleAbort() {
        if (!settled) {
          const message: MainToWorkerMessage = {
            type: 'parse-cancel',
            version: OFFICE_PARSER_PROTOCOL_VERSION,
            taskId,
            documentSessionId,
            sequence: 1,
          };
          worker.postMessage(message);
        }
        fail(createParseAbortError());
      }
      function handleWorkerError() {
        fail(
          createRuntimeError(
            parseStarted ? 'WORKER_RUNTIME_CRASH' : 'WORKER_STARTUP_FAILED',
            parseStarted
              ? 'Office 解析 Worker 运行异常'
              : 'Office 解析 Worker 加载失败',
            !parseStarted,
          ),
        );
      }
      const processMessage = async (message: WorkerToMainMessage) => {
        if (settled) return;
        if (message.version !== OFFICE_PARSER_PROTOCOL_VERSION) {
          fail(
            createRuntimeError(
              'WORKER_STARTUP_FAILED',
              'Office 解析 Worker 协议版本不匹配',
              true,
            ),
          );
          return;
        }
        if (message.type === 'worker-ready') {
          if (parseStarted || settled || signal.aborted) return;
          if (readyTimeoutId !== undefined) clearTimeout(readyTimeoutId);
          const startMessage: MainToWorkerMessage = {
            type: 'parse-start',
            version: OFFICE_PARSER_PROTOCOL_VERSION,
            taskId,
            documentSessionId,
            kind,
            fileName: file.name,
            sequence: 0,
            file,
          };
          parseStarted = true;
          worker.postMessage(startMessage);
          return;
        }
        if (
          message.taskId !== taskId ||
          message.documentSessionId !== documentSessionId
        ) {
          return;
        }
        if (message.type === 'parse-progress') {
          sink.progress(message.progress);
          return;
        }
        if (message.type === 'parse-resource') {
          await sink.resource(message.resource);
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-sheet') {
          await sink.sheet(message.sheetIndex, message.revision, message.sheet);
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-presentation-meta') {
          await sink.presentationMetadata(message.metadata);
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-slide') {
          await sink.slide(message.slideIndex, message.slide);
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-document-meta') {
          await sink.documentMetadata(message.metadata);
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-document-blocks') {
          await sink.documentBlocks(message.startIndex, message.blocks);
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-spreadsheet-meta') {
          await sink.spreadsheetMetadata(message.metadata);
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-docx-meta') {
          await sink.docxMetadata(message.metadata);
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-docx-blocks') {
          await sink.docxBlocks(message.startIndex, message.blocks);
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-docx-pages') {
          await sink.docxPages(message.startIndex, message.pages);
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-complete') {
          await sink.complete(message.warnings);
          finish();
          return;
        }
        if (message.type === 'parse-cancelled') {
          fail(createParseAbortError());
          return;
        }
        if (message.type === 'parse-error') {
          fail(deserializeParseError(message.error));
        }
      };
      function handleMessage(event: MessageEvent<WorkerToMainMessage>) {
        void processMessage(event.data).catch(fail);
      }

      this.stopActive = handleAbort;
      signal.addEventListener('abort', handleAbort, { once: true });
      worker.addEventListener('message', handleMessage);
      worker.addEventListener('error', handleWorkerError);
      readyTimeoutId = setTimeout(() => {
        fail(
          createRuntimeError(
            'WORKER_STARTUP_FAILED',
            'Office 解析 Worker 就绪超时',
            true,
          ),
        );
      }, Math.max(0, this.readyTimeoutMs));
      if (signal.aborted) handleAbort();
    });
  }

  dispose() {
    this.stopActive?.();
    this.worker?.terminate();
    this.worker = undefined;
    this.stopActive = undefined;
  }
}

/** 创建带稳定错误码的 Worker 配置错误。 */
export function createWorkerConfigurationError(
  code: 'WORKER_FORMAT_NOT_READY' | 'WORKER_UNAVAILABLE',
  message: string,
) {
  return createRuntimeError(code, message, false);
}
