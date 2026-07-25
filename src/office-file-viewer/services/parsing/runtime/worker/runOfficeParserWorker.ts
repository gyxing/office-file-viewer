import { parseDocCore } from '../../../doc/parseDocCore';
import { parsePptCore } from '../../../ppt/parsePptCore';
import { parseXlsCore } from '../../../xls/parseXlsCore';
import { serializeParseError } from '../../protocol/errors';
import type {
  MainToWorkerMessage,
  PortableResource,
  WorkerToMainMessage,
} from '../../protocol/messages';
import { OFFICE_PARSER_PROTOCOL_VERSION } from '../../protocol/version';
import { createParseAbortError } from '../types';

/** 描述 WorkerMessageEvent 在解析运行时中的数据结构。 */
type WorkerMessageEvent = {
  /** WorkerMessageEvent 当前步骤需要处理的原始或标准化数据。 */
  data: MainToWorkerMessage;
};

/** 描述 ParserWorkerScope 在解析运行时中的数据结构。 */
type ParserWorkerScope = {
  /** 执行 ParserWorkerScope 的 postMessage 操作。 */
  postMessage(message: WorkerToMainMessage, transfer?: Transferable[]): void;
  /** 执行 ParserWorkerScope 的 addEventListener 操作。 */
  addEventListener(
    type: 'message',
    listener: (event: WorkerMessageEvent) => void,
  ): void;
};

/** 执行 `resourceTransferList` 封装的解析运行时处理步骤。 */
function resourceTransferList(resource: PortableResource): Transferable[] {
  return resource.encoding === 'text' ? [] : [resource.buffer];
}

/** 在独立 Worker 中处理单个 Office 解析任务。 */
export function runOfficeParserWorker(scope: ParserWorkerScope) {
  let activeTaskId: string | undefined;
  let cancelled = false;
  let nextSequence = 1;
  const ackWaiters = new Map<number, () => void>();

  const post = (message: WorkerToMainMessage, transfer: Transferable[] = []) =>
    scope.postMessage(message, transfer);

  const waitForAck = (sequence: number) =>
    new Promise<void>((resolve) => {
      ackWaiters.set(sequence, resolve);
    });

  async function sendSequenced(
    message:
      | Extract<
          WorkerToMainMessage,
          {
            /** 用于区分 当前结构 不同结构分支的类型标识。 */
            type: 'parse-resource';
          }
        >
      | Extract<
          WorkerToMainMessage,
          {
            /** 用于区分 当前结构 不同结构分支的类型标识。 */
            type: 'parse-sheet';
          }
        >
      | Extract<
          WorkerToMainMessage,
          {
            /** 用于区分 当前结构 不同结构分支的类型标识。 */
            type: 'parse-presentation-meta';
          }
        >
      | Extract<
          WorkerToMainMessage,
          {
            /** 用于区分 当前结构 不同结构分支的类型标识。 */
            type: 'parse-slide';
          }
        >
      | Extract<
          WorkerToMainMessage,
          {
            /** 用于区分 当前结构 不同结构分支的类型标识。 */
            type: 'parse-document-meta';
          }
        >
      | Extract<
          WorkerToMainMessage,
          {
            /** 用于区分 当前结构 不同结构分支的类型标识。 */
            type: 'parse-document-blocks';
          }
        >,
    transfer: Transferable[] = [],
  ) {
    post(message, transfer);
    await waitForAck(message.sequence);
    if (cancelled) throw createParseAbortError();
  }

  async function parseXls(
    message: Extract<
      MainToWorkerMessage,
      {
        /** 用于区分 当前结构 不同结构分支的类型标识。 */ type: 'parse-start';
      }
    >,
  ) {
    const { taskId } = message;
    try {
      const result = await parseXlsCore(message.buffer, {
        checkpoint: async (progress) => {
          if (cancelled) throw createParseAbortError();
          if (progress) {
            post({
              type: 'parse-progress',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              progress,
            });
          }
        },
        output: {
          resource: async (resource) => {
            const sequence = nextSequence++;
            await sendSequenced(
              {
                type: 'parse-resource',
                version: OFFICE_PARSER_PROTOCOL_VERSION,
                taskId,
                sequence,
                resource,
              },
              resourceTransferList(resource),
            );
          },
          sheet: async (sheetIndex, revision, sheet) => {
            await sendSequenced({
              type: 'parse-sheet',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              sequence: nextSequence++,
              sheetIndex,
              revision,
              sheet,
            });
          },
        },
      });
      post({
        type: 'parse-complete',
        version: OFFICE_PARSER_PROTOCOL_VERSION,
        taskId,
        warnings: result.workbook.warnings,
      });
    } catch (error) {
      if (
        cancelled ||
        (error instanceof Error && error.name === 'AbortError')
      ) {
        post({
          type: 'parse-cancelled',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId,
        });
      } else {
        post({
          type: 'parse-error',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId,
          error: serializeParseError(error, { format: 'xls' }),
        });
      }
    }
  }

  async function parsePpt(
    message: Extract<
      MainToWorkerMessage,
      {
        /** 用于区分 当前结构 不同结构分支的类型标识。 */ type: 'parse-start';
      }
    >,
  ) {
    const { taskId } = message;
    try {
      await parsePptCore(message.buffer, {
        checkpoint: async (progress) => {
          if (cancelled) throw createParseAbortError();
          if (progress) {
            post({
              type: 'parse-progress',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              progress,
            });
          }
        },
        output: {
          resource: async (resource) => {
            const sequence = nextSequence++;
            await sendSequenced(
              {
                type: 'parse-resource',
                version: OFFICE_PARSER_PROTOCOL_VERSION,
                taskId,
                sequence,
                resource,
              },
              resourceTransferList(resource),
            );
          },
          presentationMetadata: async (metadata) => {
            await sendSequenced({
              type: 'parse-presentation-meta',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              sequence: nextSequence++,
              metadata,
            });
          },
          slide: async (slideIndex, slide) => {
            await sendSequenced({
              type: 'parse-slide',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              sequence: nextSequence++,
              slideIndex,
              slide,
            });
          },
        },
      });
      post({
        type: 'parse-complete',
        version: OFFICE_PARSER_PROTOCOL_VERSION,
        taskId,
      });
    } catch (error) {
      if (
        cancelled ||
        (error instanceof Error && error.name === 'AbortError')
      ) {
        post({
          type: 'parse-cancelled',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId,
        });
      } else {
        post({
          type: 'parse-error',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId,
          error: serializeParseError(error, { format: 'ppt' }),
        });
      }
    }
  }

  async function parseDoc(
    message: Extract<
      MainToWorkerMessage,
      {
        /** 用于区分 当前结构 不同结构分支的类型标识。 */ type: 'parse-start';
      }
    >,
  ) {
    const { taskId } = message;
    try {
      await parseDocCore(message.buffer, {
        fileName: message.fileName,
        checkpoint: async (progress) => {
          if (cancelled) throw createParseAbortError();
          if (progress) {
            post({
              type: 'parse-progress',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              progress,
            });
          }
        },
        output: {
          resource: async (resource) => {
            const sequence = nextSequence++;
            await sendSequenced(
              {
                type: 'parse-resource',
                version: OFFICE_PARSER_PROTOCOL_VERSION,
                taskId,
                sequence,
                resource,
              },
              resourceTransferList(resource),
            );
          },
          documentMetadata: async (metadata) => {
            await sendSequenced({
              type: 'parse-document-meta',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              sequence: nextSequence++,
              metadata,
            });
          },
          documentBlocks: async (startIndex, blocks) => {
            await sendSequenced({
              type: 'parse-document-blocks',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              sequence: nextSequence++,
              startIndex,
              blocks,
            });
          },
        },
      });
      post({
        type: 'parse-complete',
        version: OFFICE_PARSER_PROTOCOL_VERSION,
        taskId,
      });
    } catch (error) {
      if (
        cancelled ||
        (error instanceof Error && error.name === 'AbortError')
      ) {
        post({
          type: 'parse-cancelled',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId,
        });
      } else {
        post({
          type: 'parse-error',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId,
          error: serializeParseError(error, { format: 'doc' }),
        });
      }
    }
  }

  scope.addEventListener('message', (event) => {
    const message = event.data;
    if (message.version !== OFFICE_PARSER_PROTOCOL_VERSION) return;
    if (message.type === 'parse-ack') {
      if (message.taskId !== activeTaskId) return;
      const resolve = ackWaiters.get(message.sequence);
      ackWaiters.delete(message.sequence);
      resolve?.();
      return;
    }
    if (message.type === 'parse-cancel') {
      if (message.taskId === activeTaskId) cancelled = true;
      return;
    }
    if (activeTaskId) {
      post({
        type: 'parse-error',
        version: OFFICE_PARSER_PROTOCOL_VERSION,
        taskId: message.taskId,
        error: {
          code: 'WORKER_BUSY',
          message: '解析 Worker 正在处理其他任务',
          format: message.kind,
          recoverable: true,
        },
      });
      return;
    }
    activeTaskId = message.taskId;
    cancelled = false;
    if (message.kind === 'ppt') void parsePpt(message);
    else if (message.kind === 'doc') void parseDoc(message);
    else void parseXls(message);
  });

  post({
    type: 'worker-ready',
    version: OFFICE_PARSER_PROTOCOL_VERSION,
  });
}
