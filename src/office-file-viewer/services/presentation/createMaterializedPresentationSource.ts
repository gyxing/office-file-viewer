import { PresentationSearchProvider } from './PresentationSearchProvider';
import type {
  PresentationSlideDescriptor,
  PresentationSource,
  PresentationSourceSnapshot,
} from './PresentationSource';
import { throwIfPresentationAborted } from './PresentationSource';
import { profileMaterializedPresentation } from './presentationPerformance';
import type { PresentationDocument } from './types';

/** 将现有完整 PresentationDocument 适配为统一的按页读取接口。 */
export function createMaterializedPresentationSource(
  document: PresentationDocument,
): PresentationSource {
  const slides: PresentationSlideDescriptor[] = document.slides.map(
    (slide) => ({
      id: slide.id,
      index: slide.index,
      hidden: Boolean(slide.hidden),
      hasSpeakerNotes: Boolean(slide.speakerNotes),
      annotationCount: slide.annotations?.length ?? 0,
      estimatedElementCount: slide.elements.length,
      revision: 1,
      status: 'ready',
    }),
  );
  const snapshot: PresentationSourceSnapshot = {
    revision: 1,
    width: document.width,
    height: document.height,
    theme: document.theme,
    slideCount: slides.length,
    slides,
    warnings: document.warnings,
    performance: profileMaterializedPresentation(document.slides),
  };
  let disposed = false;

  const ensureAvailable = (index?: number) => {
    if (disposed) throw new Error('演示文稿数据源已释放');
    if (
      index !== undefined &&
      (!Number.isInteger(index) || index < 0 || index >= document.slides.length)
    ) {
      throw new RangeError(`幻灯片索引超出范围：${index}`);
    }
  };

  let searchProvider: PresentationSearchProvider;
  const source: PresentationSource = {
    get searchProvider() {
      return searchProvider;
    },
    getSnapshot() {
      return snapshot;
    },
    subscribe() {
      return () => undefined;
    },
    async getSlide(index, signal) {
      throwIfPresentationAborted(signal);
      ensureAvailable(index);
      return document.slides[index];
    },
    async getSpeakerNotes(index, signal) {
      throwIfPresentationAborted(signal);
      ensureAvailable(index);
      return document.slides[index].speakerNotes;
    },
    async getAnnotations(index, signal) {
      throwIfPresentationAborted(signal);
      ensureAvailable(index);
      return document.slides[index].annotations ?? [];
    },
    async ensureRange(start, end, signal) {
      throwIfPresentationAborted(signal);
      ensureAvailable();
      if (start > end || end < 0 || start >= document.slides.length) return;
    },
    retainRange() {
      ensureAvailable();
      return () => undefined;
    },
    retry() {
      ensureAvailable();
    },
    async dispose() {
      disposed = true;
    },
  };
  searchProvider = new PresentationSearchProvider(source);
  return source;
}
