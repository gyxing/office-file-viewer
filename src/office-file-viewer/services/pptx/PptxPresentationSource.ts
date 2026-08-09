import type { OfficeArchiveResourcePolicy } from '../../shared/resource/OfficeResourcePolicy';
import { createContentStore } from '../content-store/createContentStore';
import type { OfficeContentStore } from '../content-store/types';
import type { OfficeSourcePreviewFactory } from '../parsing/formatParserRegistry';
import {
  getPresentationSlideWeight,
  type PresentationPerformanceProfile,
} from '../presentation/presentationPerformance';
import {
  createPresentationAbortError,
  throwIfPresentationAborted,
} from '../presentation/PresentationSource';
import type {
  PresentationSlideDescriptor,
  PresentationSource,
  PresentationSourceSnapshot,
  SlideModel,
  SpeakerNotesModel,
} from '../presentation/types';
import { disposeDocumentSession } from '../session';
import { loadPptxSlide } from './loadPptxSlide';
import { parsePptxSpeakerNotes } from './parseSpeakerNotes';
import type {
  PptxPackageContext,
  PptxSlideDescriptor,
} from './PptxPackageContext';
import {
  profilePptxArchive,
  readPptxStructure,
  type ProfiledPptxArchive,
} from './readPptxStructure';

/** 幻灯片按需存储中的位置和元素统计信息。 */
type SlideStoreMeta = {
  /** 在所属集合中的零基索引。 */
  index: number;
};

function waitForSharedResult<T>(
  promise: Promise<T>,
  signal?: AbortSignal,
): Promise<T> {
  if (!signal) return promise;
  if (signal.aborted) return Promise.reject(createPresentationAbortError());
  return new Promise<T>((resolve, reject) => {
    const abort = () => reject(createPresentationAbortError());
    signal.addEventListener('abort', abort, { once: true });
    promise.then(
      (value) => {
        signal.removeEventListener('abort', abort);
        resolve(value);
      },
      (error) => {
        signal.removeEventListener('abort', abort);
        reject(error);
      },
    );
  });
}

/** 提供 PPTX 当前页优先、邻近预取和可重试缓存能力。 */
export class PptxPresentationSource implements PresentationSource {
  private readonly listeners = new Set<() => void>();
  private descriptors: PptxSlideDescriptor[];
  private readonly slideStore: OfficeContentStore<SlideStoreMeta, SlideModel>;
  private readonly slideRequests = new Map<number, Promise<SlideModel>>();
  private readonly notesRequests = new Map<
    number,
    Promise<SpeakerNotesModel | undefined>
  >();
  private readonly notes = new Map<number, SpeakerNotesModel | undefined>();
  private revision = 1;
  private disposed = false;
  private disposePromise?: Promise<void>;
  private performance: PresentationPerformanceProfile;
  private snapshot: PresentationSourceSnapshot;

  constructor(
    private readonly context: PptxPackageContext,
    performance: PresentationPerformanceProfile,
  ) {
    this.descriptors = context.descriptors.map((descriptor) => ({
      ...descriptor,
    }));
    this.performance = { ...performance, slideMode: 'lazy' };
    this.snapshot = this.createSnapshot();
    this.slideStore = createContentStore({
      sessionId: context.sessionId,
      namespace: 'pptx-slides',
      maxMemoryBytes: 48 * 1024 * 1024,
      estimateSize: (slide) =>
        2048 + slide.elements.length * 512 + getPresentationSlideWeight(slide),
    });
  }

  private ensureAvailable(index?: number) {
    if (this.disposed) throw new Error('PPTX 数据源已释放');
    if (
      index !== undefined &&
      (!Number.isInteger(index) ||
        index < 0 ||
        index >= this.descriptors.length)
    ) {
      throw new RangeError(`幻灯片索引超出范围：${index}`);
    }
  }

  private publish() {
    this.revision += 1;
    this.snapshot = this.createSnapshot();
    this.listeners.forEach((listener) => listener());
  }

  private createSnapshot(): PresentationSourceSnapshot {
    return {
      revision: this.revision,
      width: this.context.width,
      height: this.context.height,
      theme: this.context.theme,
      slideCount: this.descriptors.length,
      slides: this.descriptors,
      performance: { ...this.performance },
    };
  }

  private updateDescriptor(
    index: number,
    patch: Partial<PresentationSlideDescriptor>,
  ) {
    const current = this.descriptors[index];
    const nextDescriptors = this.descriptors.slice();
    nextDescriptors[index] = {
      ...current,
      ...patch,
      revision: current.revision + 1,
    };
    // 已发布快照继续复用旧数组，新修订只复制描述符索引，保持快照不可变。
    this.descriptors = nextDescriptors;
    this.publish();
  }

  getSnapshot(): PresentationSourceSnapshot {
    return this.snapshot;
  }

  subscribe(listener: () => void) {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  async getSlide(index: number, signal?: AbortSignal) {
    this.ensureAvailable(index);
    throwIfPresentationAborted(signal);
    const key = `slide:${index}`;
    const cached = await this.slideStore.get(key, signal);
    if (cached?.value) return cached.value;

    let request = this.slideRequests.get(index);
    if (!request) {
      const descriptor = this.descriptors[index];
      request = loadPptxSlide(this.context, descriptor).then(
        async (slide) => {
          const nextRevision = descriptor.revision + 1;
          await this.slideStore.put({
            key,
            revision: nextRevision,
            meta: { index },
            value: slide,
            updatedAt: Date.now(),
          });
          this.performance = {
            ...this.performance,
            totalElementWeight:
              this.performance.totalElementWeight +
              getPresentationSlideWeight(slide),
          };
          this.updateDescriptor(index, {
            status: 'ready',
            errorMessage: undefined,
            estimatedElementCount: slide.elements.length,
          });
          return slide;
        },
        (error) => {
          this.updateDescriptor(index, {
            status: 'error',
            errorMessage:
              error instanceof Error ? error.message : '幻灯片加载失败',
          });
          throw error;
        },
      );
      this.slideRequests.set(index, request);
      void request.then(
        () => this.slideRequests.delete(index),
        () => this.slideRequests.delete(index),
      );
    }
    return waitForSharedResult(request, signal);
  }

  async getSpeakerNotes(index: number, signal?: AbortSignal) {
    this.ensureAvailable(index);
    throwIfPresentationAborted(signal);
    if (this.notes.has(index)) return this.notes.get(index);
    let request = this.notesRequests.get(index);
    if (!request) {
      const descriptor = this.descriptors[index];
      request = (async () => {
        if (!descriptor.notesPath) return undefined;
        const xml = await this.context.reader.readText(descriptor.notesPath!);
        this.context.packageState.entries.set(descriptor.notesPath!, xml);
        const value = parsePptxSpeakerNotes(
          this.context.packageState.entries,
          this.context.packageState.relationships,
          descriptor.relsPath,
        );
        this.notes.set(index, value);
        return value;
      })();
      this.notesRequests.set(index, request);
      void request.then(
        () => this.notesRequests.delete(index),
        () => this.notesRequests.delete(index),
      );
    }
    return waitForSharedResult(request, signal);
  }

  async ensureRange(start: number, end: number, signal?: AbortSignal) {
    this.ensureAvailable();
    const normalizedStart = Math.max(0, start);
    const normalizedEnd = Math.min(this.descriptors.length - 1, end);
    if (normalizedStart > normalizedEnd) return;
    await Promise.all(
      Array.from({ length: normalizedEnd - normalizedStart + 1 }, (_, offset) =>
        this.getSlide(normalizedStart + offset, signal),
      ),
    );
  }

  retainRange(start: number, end: number) {
    this.ensureAvailable();
    const keys = this.descriptors
      .slice(Math.max(0, start), Math.min(this.descriptors.length, end + 1))
      .map((_, offset) => `slide:${Math.max(0, start) + offset}`);
    return this.slideStore.pin(keys);
  }

  retry(index: number) {
    this.ensureAvailable(index);
    if (this.descriptors[index].status !== 'error') return;
    void this.slideStore.delete(`slide:${index}`);
    this.updateDescriptor(index, {
      status: 'estimated',
      errorMessage: undefined,
    });
  }

  dispose() {
    if (this.disposePromise) return this.disposePromise;
    this.disposed = true;
    this.listeners.clear();
    this.disposePromise = Promise.allSettled([
      this.slideStore.dispose(),
      this.context.reader.close(),
    ]).then(() => undefined);
    return this.disposePromise;
  }
}

/** 从已完成中央目录画像的 Reader 创建大型 PPTX 数据源。 */
export async function createPptxPresentationSourceFromArchive(
  archive: ProfiledPptxArchive,
  sessionId: string,
  signal?: AbortSignal,
) {
  const context = await readPptxStructure(
    archive.reader,
    sessionId,
    archive.profile,
    signal,
  );
  return new PptxPresentationSource(context, archive.profile.performance);
}

/** 仅在 PPTX 画像命中大文件阈值时创建按页预览源。 */
export const tryCreatePptxSourcePreview: OfficeSourcePreviewFactory = async (
  file,
  { documentSession, emitProgress, emitPartial, resourcePolicy },
) => {
  emitProgress({
    stage: 'container',
    percent: 0.02,
    message: '正在读取 PPTX 包目录',
  });
  const archive = await profilePptxArchive(
    file,
    documentSession.signal,
    resourcePolicy,
  );
  if (archive.profile.performance.slideMode !== 'lazy') {
    await archive.reader.close();
    return undefined;
  }

  let source: PptxPresentationSource | undefined;
  try {
    source = await createPptxPresentationSourceFromArchive(
      archive,
      documentSession.id,
      documentSession.signal,
    );
    documentSession.register({ dispose: () => source?.dispose() });
    documentSession.transferTo(source);
    const state = {
      sessionId: documentSession.id,
      previewKind: 'pptx' as const,
      mode: 'source' as const,
      source,
      summary: source.getSnapshot(),
    };
    emitPartial(state);
    return {
      ...state,
      dispose: () => disposeDocumentSession(source),
    };
  } catch (error) {
    await (source?.dispose() ?? archive.reader.close());
    throw error;
  }
};

/** 打开并创建大型 PPTX 数据源，失败时保证 Reader 被关闭。 */
export async function createPptxPresentationSource(
  file: File,
  sessionId: string,
  signal?: AbortSignal,
  resourcePolicy?: OfficeArchiveResourcePolicy,
) {
  const archive = await profilePptxArchive(file, signal, resourcePolicy);
  try {
    return await createPptxPresentationSourceFromArchive(
      archive,
      sessionId,
      signal,
    );
  } catch (error) {
    await archive.reader.close();
    throw error;
  }
}
