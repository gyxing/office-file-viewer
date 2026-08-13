import { parseDocCore, type DocCoreContext } from '../../../doc/parseDocCore';
import { parseDocRandomAccess } from '../../../doc/parseDocRandomAccess';
import { OFFICE_LARGE_FILE_THRESHOLDS } from '../../../performance/officePerformanceThresholds';
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
import { runOoxmlMaterializedParser } from './runOoxmlMaterializedParser';
import { WorkerSourceHost } from './WorkerSourceHost';

/** 解析 Worker 接收消息时使用的最小事件结构。 */
type WorkerMessageEvent = {
  /** Worker 收到的解析协议消息。 */
  data: MainToWorkerMessage;
};

/** 解析 Worker 依赖的消息发送和监听接口。 */
type ParserWorkerScope = {
  /** 向主线程发送解析 Worker 协议消息。 */
  postMessage(message: WorkerToMainMessage, transfer?: Transferable[]): void;
  /** 监听主线程发送给解析 Worker 的消息。 */
  addEventListener(
    type: 'message',
    listener: (event: WorkerMessageEvent) => void,
  ): void;
};

function resourceTransferList(resource: PortableResource): Transferable[] {
  return resource.encoding === 'text' ? [] : [resource.buffer];
}

/** 在独立 Worker 中处理单个 Office 解析任务。 */
export function runOfficeParserWorker(scope: ParserWorkerScope) {
  let activeTaskId: string | undefined;
  let activeDocumentSessionId: string | undefined;
  let cancelled = false;
  let activeAbortController = new AbortController();
  let nextSequence = 1;
  const ackWaiters = new Map<number, () => void>();

  const post = (message: WorkerToMainMessage, transfer: Transferable[] = []) =>
    scope.postMessage(message, transfer);
  const sourceHost = new WorkerSourceHost(post);

  const waitForAck = (sequence: number) =>
    new Promise<void>((resolve) => {
      ackWaiters.set(sequence, resolve);
    });

  async function sendSequenced(
    message:
      | Extract<
          WorkerToMainMessage,
          {
            /** 固定为 `parse-resource`，用于区分联合类型分支。 */
            type: 'parse-resource';
          }
        >
      | Extract<
          WorkerToMainMessage,
          {
            /** 固定为 `parse-spreadsheet-meta`，用于区分联合类型分支。 */
            type: 'parse-spreadsheet-meta';
          }
        >
      | Extract<
          WorkerToMainMessage,
          {
            /** 固定为 `parse-sheet`，用于区分联合类型分支。 */
            type: 'parse-sheet';
          }
        >
      | Extract<
          WorkerToMainMessage,
          {
            /** 固定为 `parse-presentation-meta`，用于区分联合类型分支。 */
            type: 'parse-presentation-meta';
          }
        >
      | Extract<
          WorkerToMainMessage,
          {
            /** 固定为 `parse-slide`，用于区分联合类型分支。 */
            type: 'parse-slide';
          }
        >
      | Extract<
          WorkerToMainMessage,
          {
            /** 固定为 `parse-document-meta`，用于区分联合类型分支。 */
            type: 'parse-document-meta';
          }
        >
      | Extract<
          WorkerToMainMessage,
          {
            /** 固定为 `parse-document-blocks`，用于区分联合类型分支。 */
            type: 'parse-document-blocks';
          }
        >
      | Extract<
          WorkerToMainMessage,
          {
            /** 固定为 `parse-docx-meta`，用于区分联合类型分支。 */
            type: 'parse-docx-meta';
          }
        >
      | Extract<
          WorkerToMainMessage,
          {
            /** 固定为 `parse-docx-blocks`，用于区分联合类型分支。 */
            type: 'parse-docx-blocks';
          }
        >
      | Extract<
          WorkerToMainMessage,
          {
            /** 固定为 `parse-docx-pages`，用于区分联合类型分支。 */
            type: 'parse-docx-pages';
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
        /** 固定为 `parse-start`，用于区分联合类型分支。 */
        type: 'parse-start';
      }
    >,
  ) {
    const { documentSessionId, taskId } = message;
    try {
      const input = await message.file.arrayBuffer();
      if (cancelled) throw createParseAbortError();
      const result = await parseXlsCore(input, {
        checkpoint: async (progress) => {
          if (cancelled) throw createParseAbortError();
          if (progress) {
            post({
              type: 'parse-progress',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              documentSessionId,
              sequence: nextSequence++,
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
                documentSessionId,
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
              documentSessionId,
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
        documentSessionId,
        sequence: nextSequence++,
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
          documentSessionId,
          sequence: nextSequence++,
        });
      } else {
        post({
          type: 'parse-error',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId,
          documentSessionId,
          sequence: nextSequence++,
          error: serializeParseError(error, { format: 'xls' }),
        });
      }
    }
  }

  async function parsePpt(
    message: Extract<
      MainToWorkerMessage,
      {
        /** 固定为 `parse-start`，用于区分联合类型分支。 */
        type: 'parse-start';
      }
    >,
  ) {
    const { documentSessionId, taskId } = message;
    try {
      const input = await message.file.arrayBuffer();
      if (cancelled) throw createParseAbortError();
      await parsePptCore(input, {
        checkpoint: async (progress) => {
          if (cancelled) throw createParseAbortError();
          if (progress) {
            post({
              type: 'parse-progress',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              documentSessionId,
              sequence: nextSequence++,
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
                documentSessionId,
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
              documentSessionId,
              sequence: nextSequence++,
              metadata,
            });
          },
          slide: async (slideIndex, slide) => {
            await sendSequenced({
              type: 'parse-slide',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              documentSessionId,
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
        documentSessionId,
        sequence: nextSequence++,
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
          documentSessionId,
          sequence: nextSequence++,
        });
      } else {
        post({
          type: 'parse-error',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId,
          documentSessionId,
          sequence: nextSequence++,
          error: serializeParseError(error, { format: 'ppt' }),
        });
      }
    }
  }

  async function parseDoc(
    message: Extract<
      MainToWorkerMessage,
      {
        /** 固定为 `parse-start`，用于区分联合类型分支。 */
        type: 'parse-start';
      }
    >,
  ) {
    const { documentSessionId, taskId } = message;
    try {
      const parseContext: DocCoreContext = {
        fileName: message.fileName,
        checkpoint: async (progress) => {
          if (cancelled) throw createParseAbortError();
          if (progress) {
            post({
              type: 'parse-progress',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              documentSessionId,
              sequence: nextSequence++,
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
                documentSessionId,
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
              documentSessionId,
              sequence: nextSequence++,
              metadata,
            });
          },
          documentBlocks: async (startIndex, blocks) => {
            await sendSequenced({
              type: 'parse-document-blocks',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              documentSessionId,
              sequence: nextSequence++,
              startIndex,
              blocks,
            });
          },
        },
      };
      if (message.file.size >= OFFICE_LARGE_FILE_THRESHOLDS.cfbFileBytes) {
        await parseDocRandomAccess(
          message.file,
          parseContext,
          activeAbortController.signal,
        );
      } else {
        const input = await message.file.arrayBuffer();
        if (cancelled) throw createParseAbortError();
        await parseDocCore(input, parseContext);
      }
      post({
        type: 'parse-complete',
        version: OFFICE_PARSER_PROTOCOL_VERSION,
        taskId,
        documentSessionId,
        sequence: nextSequence++,
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
          documentSessionId,
          sequence: nextSequence++,
        });
      } else {
        post({
          type: 'parse-error',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId,
          documentSessionId,
          sequence: nextSequence++,
          error: serializeParseError(error, { format: 'doc' }),
        });
      }
    }
  }

  async function parseOoxml(
    message: Extract<
      MainToWorkerMessage,
      {
        /** 固定为 `parse-start`，用于区分联合类型分支。 */
        type: 'parse-start';
      }
    >,
  ) {
    const { documentSessionId, kind, taskId } = message;
    if (kind !== 'docx' && kind !== 'xlsx' && kind !== 'pptx') return;
    try {
      const warnings = await runOoxmlMaterializedParser(
        message.file,
        kind,
        activeAbortController.signal,
        {
          progress: (progress) => {
            post({
              type: 'parse-progress',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              documentSessionId,
              sequence: nextSequence++,
              progress,
            });
          },
          docxMetadata: (metadata) =>
            sendSequenced({
              type: 'parse-docx-meta',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              documentSessionId,
              sequence: nextSequence++,
              metadata,
            }),
          docxBlocks: (startIndex, blocks) =>
            sendSequenced({
              type: 'parse-docx-blocks',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              documentSessionId,
              sequence: nextSequence++,
              startIndex,
              blocks,
            }),
          docxPages: (startIndex, pages) =>
            sendSequenced({
              type: 'parse-docx-pages',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              documentSessionId,
              sequence: nextSequence++,
              startIndex,
              pages,
            }),
          spreadsheetMetadata: (metadata) =>
            sendSequenced({
              type: 'parse-spreadsheet-meta',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              documentSessionId,
              sequence: nextSequence++,
              metadata,
            }),
          sheet: (sheetIndex, sheet) =>
            sendSequenced({
              type: 'parse-sheet',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              documentSessionId,
              sequence: nextSequence++,
              sheetIndex,
              revision: 1,
              sheet,
            }),
          presentationMetadata: (metadata) =>
            sendSequenced({
              type: 'parse-presentation-meta',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              documentSessionId,
              sequence: nextSequence++,
              metadata,
            }),
          slide: (slideIndex, slide) =>
            sendSequenced({
              type: 'parse-slide',
              version: OFFICE_PARSER_PROTOCOL_VERSION,
              taskId,
              documentSessionId,
              sequence: nextSequence++,
              slideIndex,
              slide,
            }),
        },
      );
      post({
        type: 'parse-complete',
        version: OFFICE_PARSER_PROTOCOL_VERSION,
        taskId,
        documentSessionId,
        sequence: nextSequence++,
        warnings,
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
          documentSessionId,
          sequence: nextSequence++,
        });
      } else {
        post({
          type: 'parse-error',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId,
          documentSessionId,
          sequence: nextSequence++,
          error: serializeParseError(error, { format: kind }),
        });
      }
    }
  }

  scope.addEventListener('message', (event) => {
    const message = event.data;
    if (message.version !== OFFICE_PARSER_PROTOCOL_VERSION) return;
    if (sourceHost.handle(message)) return;
    if (message.type === 'parse-ack') {
      if (
        message.taskId !== activeTaskId ||
        message.documentSessionId !== activeDocumentSessionId
      ) {
        return;
      }
      const resolve = ackWaiters.get(message.sequence);
      ackWaiters.delete(message.sequence);
      resolve?.();
      return;
    }
    if (message.type === 'parse-cancel') {
      if (
        message.taskId === activeTaskId &&
        message.documentSessionId === activeDocumentSessionId
      ) {
        cancelled = true;
        activeAbortController.abort();
        ackWaiters.forEach((resolve) => resolve());
        ackWaiters.clear();
      }
      return;
    }
    if (message.type !== 'parse-start') return;
    if (activeTaskId) {
      post({
        type: 'parse-error',
        version: OFFICE_PARSER_PROTOCOL_VERSION,
        taskId: message.taskId,
        documentSessionId: message.documentSessionId,
        sequence: nextSequence++,
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
    activeDocumentSessionId = message.documentSessionId;
    cancelled = false;
    activeAbortController = new AbortController();
    if (
      message.kind === 'docx' ||
      message.kind === 'xlsx' ||
      message.kind === 'pptx'
    ) {
      void parseOoxml(message);
    } else if (message.kind === 'ppt') void parsePpt(message);
    else if (message.kind === 'doc') void parseDoc(message);
    else void parseXls(message);
  });

  post({
    type: 'worker-ready',
    version: OFFICE_PARSER_PROTOCOL_VERSION,
  });
}
