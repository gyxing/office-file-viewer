import { profileDocxArchive } from '../../../docx/docxArchiveProfile';
import type { DocxMeasuredBlock } from '../../../docx/docxPagination';
import { DocxWordPageSource } from '../../../docx/DocxWordPageSource';
import { parseDocxSource } from '../../../docx/parseDocxSource';
import {
  createPptxPresentationSourceFromArchive,
  type PptxPresentationSource,
} from '../../../pptx/PptxPresentationSource';
import { profilePptxArchive } from '../../../pptx/readPptxStructure';
import type { OfficeResourceSource } from '../../../resource-store/types';
import type { OfficeSearchQuery } from '../../../search/types';
import type { SpreadsheetRange } from '../../../spreadsheet/types';
import { profileXlsxArchive } from '../../../xlsx/readXlsxStructure';
import {
  createXlsxSpreadsheetSourceFromArchive,
  type XlsxSpreadsheetSource,
} from '../../../xlsx/XlsxSpreadsheetSource';
import { serializeParseError } from '../../protocol/errors';
import type {
  MainToWorkerMessage,
  WorkerSourceKind,
  WorkerSourceState,
  WorkerToMainMessage,
} from '../../protocol/messages';
import { OFFICE_PARSER_PROTOCOL_VERSION } from '../../protocol/version';

/** Source Host 发送消息所需的最小 Worker 接口。 */
type SourceWorkerPost = (
  message: WorkerToMainMessage,
  transfer?: Transferable[],
) => void;

/** Worker 内当前长期持有的数据源联合。 */
type ActiveWorkerSource =
  | { kind: 'docx'; source: DocxWordPageSource }
  | { kind: 'xlsx'; source: XlsxSpreadsheetSource }
  | { kind: 'pptx'; source: PptxPresentationSource };

/** Worker 懒资源注册表中保存的来源和跨线程标识。 */
type RegisteredWorkerResource = {
  /** Worker 内实际可按需加载的资源。 */
  source: Extract<OfficeResourceSource, { kind: 'lazy' }>;
  /** 跨线程返回时使用的唯一标识。 */
  workerResourceId: string;
};

function isSourceMessage(message: MainToWorkerMessage) {
  return message.type.startsWith('source-');
}

function asRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === 'object'
    ? (value as Record<string, unknown>)
    : {};
}

/** 在解析 Worker 内管理长期 OOXML Source、RPC 和懒资源。 */
export class WorkerSourceHost {
  private activeTaskId?: string;
  private activeDocumentSessionId?: string;
  private active?: ActiveWorkerSource;
  private readonly requestControllers = new Map<number, AbortController>();
  private readonly retentions = new Map<string, () => void>();
  private readonly resources = new Map<string, RegisteredWorkerResource>();
  private readonly resourceIds = new WeakMap<object, string>();
  private resourceSequence = 0;
  private unsubscribeSource?: () => void;
  private unsubscribeOutline?: () => void;
  private lifecycleController?: AbortController;
  private opened = false;
  private updateScheduled = false;
  private disposing = false;

  constructor(private readonly post: SourceWorkerPost) {}

  /** 处理 Source 协议消息；非 Source 消息返回 false 交回物化解析器。 */
  handle(message: MainToWorkerMessage) {
    if (!isSourceMessage(message)) return false;
    if (message.type === 'source-open') {
      void this.open(message);
      return true;
    }
    if (
      message.taskId !== this.activeTaskId ||
      message.documentSessionId !== this.activeDocumentSessionId
    ) {
      return true;
    }
    if (message.type === 'source-request') {
      void this.runRequest(message);
    } else if (message.type === 'source-request-cancel') {
      this.requestControllers.get(message.requestId)?.abort();
    } else if (message.type === 'source-dispose') {
      void this.dispose(message.requestId);
    }
    return true;
  }

  private async open(
    message: Extract<MainToWorkerMessage, { type: 'source-open' }>,
  ) {
    if (this.activeTaskId || this.disposing) {
      this.postRequestError(
        message.requestId,
        message.kind,
        new Error('Source Worker 正在处理其他文档'),
        message,
      );
      return;
    }
    this.activeTaskId = message.taskId;
    this.activeDocumentSessionId = message.documentSessionId;
    const controller = new AbortController();
    this.lifecycleController = controller;
    this.requestControllers.set(message.requestId, controller);
    try {
      this.postProgress(message.requestId, {
        stage: 'container',
        percent: 0.02,
        message: `正在 Worker 中读取 ${message.kind.toUpperCase()} 包目录`,
      });
      const available = await this.createSource(
        message.file,
        message.kind,
        message.resourcePolicy,
        controller.signal,
      );
      if (!available || !this.active) {
        this.post({
          type: 'source-opened',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId: message.taskId,
          documentSessionId: message.documentSessionId,
          requestId: message.requestId,
          result: { available: false },
        });
        this.resetIdentity();
        return;
      }
      this.bindSourceUpdates();
      const source = this.createPortableState();
      this.post({
        type: 'source-opened',
        version: OFFICE_PARSER_PROTOCOL_VERSION,
        taskId: message.taskId,
        documentSessionId: message.documentSessionId,
        requestId: message.requestId,
        result: { available: true, source },
      });
      this.opened = true;
      this.scheduleUpdate();
    } catch (error) {
      this.postRequestError(message.requestId, message.kind, error, message);
      await this.disposeActiveSource();
      this.resetIdentity();
    } finally {
      this.requestControllers.delete(message.requestId);
    }
  }

  private async createSource(
    file: File,
    kind: WorkerSourceKind,
    resourcePolicy: Extract<
      MainToWorkerMessage,
      { type: 'source-open' }
    >['resourcePolicy'],
    signal: AbortSignal,
  ) {
    if (kind === 'docx') {
      const archive = await profileDocxArchive(file, signal, resourcePolicy);
      if (archive.profile.mode !== 'lazy') {
        await archive.reader.close();
        return false;
      }
      const source = new DocxWordPageSource({
        sessionId: this.activeDocumentSessionId!,
        reader: archive.reader,
        signal,
      });
      this.active = { kind, source };
      let resolveMetadata: (() => void) | undefined;
      let rejectMetadata: ((error: unknown) => void) | undefined;
      const metadataReady = new Promise<void>((resolve, reject) => {
        resolveMetadata = resolve;
        rejectMetadata = reject;
      });
      void parseDocxSource(
        archive.reader,
        {
          metadata: (metadata) => {
            source.setMetadata(metadata);
            resolveMetadata?.();
            this.scheduleUpdate();
          },
          page: async (page) => {
            await source.addSourcePage(page);
            this.scheduleUpdate();
          },
          complete: (result) => {
            source.finishParsing(result);
            this.scheduleUpdate();
          },
        },
        signal,
        this.activeDocumentSessionId,
      ).catch((error) => {
        rejectMetadata?.(error);
        if (this.opened && !signal.aborted) this.postBackgroundFailure(error);
      });
      await metadataReady;
      return true;
    }
    if (kind === 'xlsx') {
      const archive = await profileXlsxArchive(file, signal, resourcePolicy);
      const created = await createXlsxSpreadsheetSourceFromArchive(
        archive,
        this.activeDocumentSessionId!,
        signal,
      );
      if (!created.profile.requiresSource) {
        await created.source.dispose();
        return false;
      }
      this.active = { kind, source: created.source };
      return true;
    }
    const archive = await profilePptxArchive(file, signal, resourcePolicy);
    if (archive.profile.performance.slideMode !== 'lazy') {
      await archive.reader.close();
      return false;
    }
    const source = await createPptxPresentationSourceFromArchive(
      archive,
      this.activeDocumentSessionId!,
      signal,
    );
    this.active = { kind, source };
    return true;
  }

  private bindSourceUpdates() {
    if (!this.active) return;
    this.unsubscribeSource = this.active.source.subscribe(() =>
      this.scheduleUpdate(),
    );
    if (this.active.kind === 'docx') {
      this.unsubscribeOutline = this.active.source.outline.subscribe(() =>
        this.scheduleUpdate(),
      );
    }
  }

  private async runRequest(
    message: Extract<MainToWorkerMessage, { type: 'source-request' }>,
  ) {
    if (!this.active || this.disposing) {
      this.postRequestError(
        message.requestId,
        this.active?.kind ?? 'docx',
        new Error('Source 尚未打开或正在释放'),
        message,
      );
      return;
    }
    const requestKind = this.active.kind;
    const controller = new AbortController();
    this.requestControllers.set(message.requestId, controller);
    try {
      const result = await this.executeOperation(
        message.operation,
        message.args,
        message.requestId,
        controller.signal,
      );
      const portable = this.externalizeValue(result);
      const transfer = this.collectTransferables(portable);
      this.post(
        {
          type: 'source-response',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId: message.taskId,
          documentSessionId: message.documentSessionId,
          requestId: message.requestId,
          result: portable,
        },
        transfer,
      );
    } catch (error) {
      this.postRequestError(message.requestId, requestKind, error, message);
    } finally {
      this.requestControllers.delete(message.requestId);
    }
  }

  private async executeOperation(
    operation: string,
    rawArgs: unknown,
    requestId: number,
    signal: AbortSignal,
  ): Promise<unknown> {
    const args = asRecord(rawArgs);
    const active = this.active!;
    if (operation === 'load-resource') {
      const resource = this.resources.get(String(args.resourceId))?.source;
      if (!resource) throw new Error('Worker 图片资源不存在或已释放');
      const blob = await resource.load(signal);
      return {
        mimeType: blob.type || resource.mimeType,
        buffer: await blob.arrayBuffer(),
      };
    }
    if (operation === 'release-retention') {
      const retentionId = String(args.retentionId);
      this.retentions.get(retentionId)?.();
      this.retentions.delete(retentionId);
      return undefined;
    }
    if (operation === 'search') {
      const provider = active.source.searchProvider;
      if (!provider) return undefined;
      const query = args.query as OfficeSearchQuery;
      await provider.search(
        query,
        (progress) => {
          if (signal.aborted) return;
          this.post({
            type: 'source-search-progress',
            version: OFFICE_PARSER_PROTOCOL_VERSION,
            taskId: this.activeTaskId!,
            documentSessionId: this.activeDocumentSessionId!,
            requestId,
            query,
            progress,
          });
        },
        signal,
      );
      return undefined;
    }
    if (active.kind === 'docx') {
      const source = active.source;
      if (operation === 'get-page') {
        return source.getPage(Number(args.index), signal);
      }
      if (operation === 'ensure-word-range') {
        return source.ensureRange(Number(args.start), Number(args.end), signal);
      }
      if (operation === 'prioritize-block') {
        return source.prioritizeBlock(String(args.blockId), signal);
      }
      if (operation === 'retain-word-range') {
        this.storeRetention(
          String(args.retentionId),
          source.retainRange(Number(args.start), Number(args.end)),
        );
        return undefined;
      }
      if (operation === 'retry-word-page') {
        source.retry(Number(args.index));
        return undefined;
      }
      if (operation === 'commit-word-measurement') {
        const batch = source.getMeasurementBatch();
        if (!batch || batch.id !== String(args.batchId)) return undefined;
        return source.commitMeasurement(
          batch,
          (args.measurements ?? []) as DocxMeasuredBlock[],
          Number(args.durationMs),
        );
      }
      if (operation === 'fail-word-measurement') {
        const batch = source.getMeasurementBatch();
        if (batch?.id === String(args.batchId)) {
          source.failMeasurement(batch, new Error(String(args.message)));
        }
        return undefined;
      }
    }
    if (active.kind === 'xlsx') {
      const source = active.source;
      const sheetId = String(args.sheetId);
      if (operation === 'ensure-sheet') {
        return source.ensureSheet(sheetId, signal);
      }
      if (operation === 'get-materialized-sheet') {
        return source.getMaterializedSheet(sheetId, signal);
      }
      if (operation === 'get-range') {
        return source.getRange(sheetId, args.range as SpreadsheetRange, signal);
      }
      if (operation === 'retain-sheet-range') {
        this.storeRetention(
          String(args.retentionId),
          source.retainRange(sheetId, args.range as SpreadsheetRange),
        );
        return undefined;
      }
      if (operation === 'retry-sheet') {
        source.retrySheet(sheetId);
        return undefined;
      }
    }
    if (active.kind === 'pptx') {
      const source = active.source;
      if (operation === 'get-slide') {
        return source.getSlide(Number(args.index), signal);
      }
      if (operation === 'get-notes') {
        return source.getSpeakerNotes(Number(args.index), signal);
      }
      if (operation === 'ensure-presentation-range') {
        return source.ensureRange(Number(args.start), Number(args.end), signal);
      }
      if (operation === 'retain-presentation-range') {
        this.storeRetention(
          String(args.retentionId),
          source.retainRange(Number(args.start), Number(args.end)),
        );
        return undefined;
      }
      if (operation === 'retry-presentation-slide') {
        source.retry(Number(args.index));
        return undefined;
      }
    }
    throw new Error(`不支持的 Worker Source 操作：${operation}`);
  }

  private storeRetention(retentionId: string, release: () => void) {
    this.retentions.get(retentionId)?.();
    this.retentions.set(retentionId, release);
  }

  private createPortableState(): WorkerSourceState {
    const active = this.active!;
    if (active.kind === 'docx') {
      const state = {
        kind: 'docx' as const,
        summary: active.source.getSummary(),
        snapshot: active.source.getSnapshot(),
        outlineItems: active.source.getOutlineItems(),
        outlineComplete: active.source.outline.getSnapshot().complete,
        performance: active.source.getPerformanceProfile(),
        measurementBatch: active.source.getMeasurementBatch(),
      };
      return this.externalizeValue(state) as WorkerSourceState;
    }
    if (active.kind === 'xlsx') {
      const snapshot = active.source.getSnapshot();
      const profileMap: Record<
        string,
        ReturnType<typeof active.source.getProfile>
      > = {};
      const layoutMap: Record<
        string,
        ReturnType<typeof active.source.getSheetLayout>
      > = {};
      snapshot.sheets.forEach((sheet) => {
        profileMap[sheet.id] = active.source.getProfile(sheet.id);
        layoutMap[sheet.id] = active.source.getSheetLayout(sheet.id);
      });
      return this.externalizeValue({
        kind: 'xlsx' as const,
        snapshot,
        profiles: profileMap,
        layouts: layoutMap,
      }) as WorkerSourceState;
    }
    return this.externalizeValue({
      kind: 'pptx' as const,
      snapshot: active.source.getSnapshot(),
    }) as WorkerSourceState;
  }

  private scheduleUpdate() {
    if (!this.opened || this.updateScheduled || !this.active) return;
    this.updateScheduled = true;
    queueMicrotask(() => {
      this.updateScheduled = false;
      if (!this.opened || !this.active || this.disposing) return;
      this.post({
        type: 'source-update',
        version: OFFICE_PARSER_PROTOCOL_VERSION,
        taskId: this.activeTaskId!,
        documentSessionId: this.activeDocumentSessionId!,
        source: this.createPortableState(),
      });
    });
  }

  private postProgress(
    requestId: number,
    progress: Extract<
      WorkerToMainMessage,
      { type: 'source-progress' }
    >['progress'],
  ) {
    this.post({
      type: 'source-progress',
      version: OFFICE_PARSER_PROTOCOL_VERSION,
      taskId: this.activeTaskId!,
      documentSessionId: this.activeDocumentSessionId!,
      requestId,
      progress,
    });
  }

  private postRequestError(
    requestId: number,
    kind: WorkerSourceKind,
    error: unknown,
    identity: Pick<
      Extract<MainToWorkerMessage, { type: 'source-open' | 'source-request' }>,
      'taskId' | 'documentSessionId'
    >,
  ) {
    this.post({
      type: 'source-error',
      version: OFFICE_PARSER_PROTOCOL_VERSION,
      taskId: identity.taskId,
      documentSessionId: identity.documentSessionId,
      requestId,
      error: serializeParseError(error, { format: kind }),
    });
  }

  private postBackgroundFailure(error: unknown) {
    if (!this.active) return;
    this.post({
      type: 'source-failed',
      version: OFFICE_PARSER_PROTOCOL_VERSION,
      taskId: this.activeTaskId!,
      documentSessionId: this.activeDocumentSessionId!,
      error: serializeParseError(error, { format: this.active.kind }),
    });
  }

  private async dispose(requestId: number) {
    if (this.disposing) return;
    this.disposing = true;
    this.lifecycleController?.abort();
    this.requestControllers.forEach((controller) => controller.abort());
    this.requestControllers.clear();
    await this.disposeActiveSource();
    this.post({
      type: 'source-disposed',
      version: OFFICE_PARSER_PROTOCOL_VERSION,
      taskId: this.activeTaskId!,
      documentSessionId: this.activeDocumentSessionId!,
      requestId,
    });
    this.resetIdentity();
    this.disposing = false;
  }

  private async disposeActiveSource() {
    this.unsubscribeSource?.();
    this.unsubscribeOutline?.();
    this.unsubscribeSource = undefined;
    this.unsubscribeOutline = undefined;
    this.retentions.forEach((release) => release());
    this.retentions.clear();
    await this.active?.source.dispose();
    this.active = undefined;
    this.resources.clear();
    this.opened = false;
  }

  private resetIdentity() {
    this.activeTaskId = undefined;
    this.activeDocumentSessionId = undefined;
    this.lifecycleController = undefined;
    this.opened = false;
    this.updateScheduled = false;
  }

  private externalizeValue(
    value: unknown,
    seen = new WeakMap<object, unknown>(),
  ): unknown {
    if (!value || typeof value !== 'object') {
      return typeof value === 'function' ? undefined : value;
    }
    if (
      value instanceof ArrayBuffer ||
      ArrayBuffer.isView(value) ||
      (typeof Blob !== 'undefined' && value instanceof Blob)
    ) {
      return value;
    }
    const resource = value as Partial<
      Extract<OfficeResourceSource, { kind: 'lazy' }>
    >;
    if (resource.kind === 'lazy' && typeof resource.load === 'function') {
      const source = value as Extract<OfficeResourceSource, { kind: 'lazy' }>;
      let workerResourceId = this.resourceIds.get(source);
      if (!workerResourceId) {
        workerResourceId = `${this.activeDocumentSessionId}:resource:${++this
          .resourceSequence}`;
        this.resourceIds.set(source, workerResourceId);
        this.resources.set(workerResourceId, { source, workerResourceId });
      }
      return {
        kind: 'lazy',
        id: source.id,
        mimeType: source.mimeType,
        size: source.size,
        workerResourceId,
      };
    }
    const existing = seen.get(value);
    if (existing) return existing;
    if (Array.isArray(value)) {
      const result: unknown[] = [];
      seen.set(value, result);
      value.forEach((item) => result.push(this.externalizeValue(item, seen)));
      return result;
    }
    const result: Record<string, unknown> = {};
    seen.set(value, result);
    Object.entries(value).forEach(([key, item]) => {
      const portable = this.externalizeValue(item, seen);
      if (portable !== undefined) result[key] = portable;
    });
    return result;
  }

  private collectTransferables(value: unknown) {
    const transfers: Transferable[] = [];
    const seen = new Set<object>();
    const visit = (item: unknown) => {
      if (!item || typeof item !== 'object') return;
      if (item instanceof ArrayBuffer) {
        transfers.push(item);
        return;
      }
      if (ArrayBuffer.isView(item)) {
        transfers.push(item.buffer);
        return;
      }
      if (seen.has(item)) return;
      seen.add(item);
      if (Array.isArray(item)) item.forEach(visit);
      else Object.values(item).forEach(visit);
    };
    visit(value);
    return [...new Set(transfers)];
  }
}
