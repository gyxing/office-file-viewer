import { createContentStore } from '../content-store/createContentStore';
import type { OfficeContentStore } from '../content-store/types';
import type { OfficeSourcePreviewFactory } from '../parsing/formatParserRegistry';
import {
  getPresentationSlideWeight,
  type PresentationPerformanceProfile,
} from '../presentation/presentationPerformance';
import { PresentationSearchProvider } from '../presentation/PresentationSearchProvider';
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
import { readPptNotes } from './document/readNotes';
import { loadPptSlide } from './loadPptSlide';
import {
  createLocalPptEditChain,
  readPptPersistObject,
} from './readPptPersistObject';
import {
  profilePptArchive,
  readPptStructure,
  type PptStructure,
  type ProfiledPptArchive,
} from './readPptStructure';

/** 幻灯片按需存储中的位置和元素统计信息。 */
type SlideStoreMeta = {
  /** 幻灯片在演示文稿中的零基索引。 */
  index: number;
};

function waitForPptResult<T>(promise: Promise<T>, signal?: AbortSignal) {
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

/** 为二进制 PPT 提供当前页优先的按需解析和备注读取。 */
export class PptPresentationSource implements PresentationSource {
  readonly searchProvider: PresentationSearchProvider;
  private readonly listeners = new Set<() => void>();
  private descriptors: PresentationSlideDescriptor[];
  private readonly store: OfficeContentStore<SlideStoreMeta, SlideModel>;
  private readonly requests = new Map<number, Promise<SlideModel>>();
  private readonly noteRequests = new Map<
    number,
    Promise<SpeakerNotesModel | undefined>
  >();
  private readonly notes = new Map<number, SpeakerNotesModel | undefined>();
  private readonly lifecycleController = new AbortController();
  private revision = 1;
  private disposed = false;
  private disposePromise?: Promise<void>;
  private performance: PresentationPerformanceProfile;
  private snapshot: PresentationSourceSnapshot;

  constructor(
    private readonly structure: PptStructure,
    performance: PresentationPerformanceProfile,
    sessionId: string,
  ) {
    this.descriptors = structure.slideDescriptors.map((descriptor) => ({
      id: `ppt-slide-${descriptor.persistId}`,
      index: descriptor.index,
      hidden: Boolean(descriptor.hidden),
      hasSpeakerNotes: structure.notesBySlideId.has(descriptor.slideId),
      annotationCount: 0,
      estimatedElementCount: 1,
      revision: 1,
      status: 'estimated',
    }));
    this.performance = {
      ...performance,
      slideMode: 'lazy',
      thumbnailMode: this.descriptors.length > 50 ? 'virtual' : 'normal',
    };
    this.store = createContentStore({
      sessionId,
      namespace: 'ppt-slides',
      maxMemoryBytes: 48 * 1024 * 1024,
      estimateSize: (slide) => 2048 + slide.elements.length * 512,
    });
    this.snapshot = this.createSnapshot();
    this.searchProvider = new PresentationSearchProvider(this);
  }

  private ensureAvailable(index?: number) {
    if (this.disposed) throw new Error('PPT 数据源已释放');
    if (
      index !== undefined &&
      (!Number.isInteger(index) ||
        index < 0 ||
        index >= this.descriptors.length)
    ) {
      throw new RangeError(`幻灯片索引超出范围：${index}`);
    }
  }

  private createSnapshot(): PresentationSourceSnapshot {
    return {
      revision: this.revision,
      width: this.structure.width,
      height: this.structure.height,
      theme: this.structure.theme,
      slideCount: this.descriptors.length,
      slides: this.descriptors,
      warnings: this.structure.warnings,
      performance: { ...this.performance },
    };
  }

  private publish() {
    this.revision += 1;
    this.snapshot = this.createSnapshot();
    this.listeners.forEach((listener) => listener());
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

  getSnapshot() {
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
    const cached = await this.store.get(key, signal);
    if (cached?.value) return cached.value;
    let request = this.requests.get(index);
    if (!request) {
      const descriptor = this.descriptors[index];
      request = loadPptSlide(
        this.structure,
        index,
        this.lifecycleController.signal,
      ).then(
        async (slide) => {
          await this.store.put({
            key,
            revision: descriptor.revision + 1,
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
            annotationCount: slide.annotations?.length ?? 0,
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
      this.requests.set(index, request);
      void request.then(
        () => this.requests.delete(index),
        () => this.requests.delete(index),
      );
    }
    return waitForPptResult(request, signal);
  }

  async getSpeakerNotes(index: number, signal?: AbortSignal) {
    this.ensureAvailable(index);
    throwIfPresentationAborted(signal);
    if (this.notes.has(index)) return this.notes.get(index);
    let request = this.noteRequests.get(index);
    if (!request) {
      request = (async () => {
        const slideDescriptor = this.structure.slideDescriptors[index];
        const descriptor = this.structure.notesBySlideId.get(
          slideDescriptor.slideId,
        );
        if (!descriptor) return undefined;
        const record = await readPptPersistObject(
          this.structure.documentStream,
          this.structure.editChain,
          descriptor.persistId,
          this.lifecycleController.signal,
        );
        if (!record) return undefined;
        return (
          await readPptNotes(
            record.bytes,
            createLocalPptEditChain(
              this.structure.editChain,
              descriptor.persistId,
            ),
            descriptor,
            this.structure.theme,
            this.structure.fonts,
            this.structure.parseContext,
          )
        )?.speakerNotes;
      })().then((result) => {
        this.notes.set(index, result);
        return result;
      });
      this.noteRequests.set(index, request);
      void request.then(
        () => this.noteRequests.delete(index),
        () => this.noteRequests.delete(index),
      );
    }
    return waitForPptResult(request, signal);
  }

  async getAnnotations(index: number, signal?: AbortSignal) {
    const slide = await this.getSlide(index, signal);
    return slide.annotations ?? [];
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
    const normalizedStart = Math.max(0, start);
    const keys = this.descriptors
      .slice(normalizedStart, Math.min(this.descriptors.length, end + 1))
      .map((_, offset) => `slide:${normalizedStart + offset}`);
    return this.store.pin(keys);
  }

  retry(index: number) {
    this.ensureAvailable(index);
    if (this.descriptors[index].status !== 'error') return;
    void this.store.delete(`slide:${index}`);
    this.updateDescriptor(index, {
      status: 'estimated',
      errorMessage: undefined,
    });
  }

  dispose() {
    if (this.disposePromise) return this.disposePromise;
    this.disposed = true;
    this.lifecycleController.abort();
    this.listeners.clear();
    const requests = [...this.requests.values(), ...this.noteRequests.values()];
    this.disposePromise = Promise.allSettled(requests)
      .then(() =>
        Promise.allSettled([
          this.store.dispose(),
          this.structure.reader.close(),
        ]),
      )
      .then(() => {
        this.structure.resources.dispose();
      });
    return this.disposePromise;
  }
}

/** 从已完成 CFB 目录画像的 Reader 创建 PPT 按页数据源。 */
export async function createPptPresentationSourceFromArchive(
  archive: ProfiledPptArchive,
  sessionId: string,
  signal?: AbortSignal,
) {
  const structure = await readPptStructure(archive, signal);
  return new PptPresentationSource(
    structure,
    archive.profile.performance,
    sessionId,
  );
}

/** 仅在 PPT 画像命中大文件阈值时创建按页预览源。 */
export const tryCreatePptSourcePreview: OfficeSourcePreviewFactory = async (
  file,
  { documentSession, emitProgress, emitPartial },
) => {
  emitProgress({
    stage: 'container',
    percent: 0.02,
    message: '正在读取 PPT 复合文档目录',
  });
  const archive = await profilePptArchive(file, documentSession.signal);
  if (archive.profile.performance.slideMode !== 'lazy') {
    await archive.reader.close();
    return undefined;
  }

  let source: PptPresentationSource | undefined;
  try {
    source = await createPptPresentationSourceFromArchive(
      archive,
      documentSession.id,
      documentSession.signal,
    );
    documentSession.register({ dispose: () => source?.dispose() });
    documentSession.transferTo(source);
    const state = {
      sessionId: documentSession.id,
      previewKind: 'ppt' as const,
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
