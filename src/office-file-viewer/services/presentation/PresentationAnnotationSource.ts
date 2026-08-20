import type { OfficeAnnotationSource } from '../annotations/AnnotationSource';
import type {
  OfficeAnnotation,
  OfficeAnnotationSourceSnapshot,
} from '../annotations/types';
import type { PresentationSource } from './PresentationSource';

/** 让共享审阅面板按幻灯片读取批注，避免大文稿一次传输全部页面模型。 */
export class PresentationAnnotationSource implements OfficeAnnotationSource {
  private readonly listeners = new Set<() => void>();
  private readonly unsubscribe: () => void;
  private revision = 1;
  private signature = '';
  private snapshot: OfficeAnnotationSourceSnapshot;

  constructor(private readonly source: PresentationSource) {
    this.snapshot = this.createSnapshot();
    this.signature = this.createSignature();
    this.unsubscribe = source.subscribe(() => this.sync());
  }

  getSnapshot() {
    return this.snapshot;
  }

  subscribe(listener: () => void) {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  async getRange(start: number, end: number, signal?: AbortSignal) {
    const { slides } = this.source.getSnapshot();
    const count = this.snapshot.count;
    const rangeStart = Math.min(count, Math.max(0, Math.trunc(start)));
    const rangeEnd = Math.min(count, Math.max(rangeStart, Math.trunc(end)));
    if (rangeStart >= rangeEnd) return [];

    let offset = 0;
    const requests: Promise<readonly OfficeAnnotation[]>[] = [];
    slides.forEach((slide, slideIndex) => {
      const nextOffset = offset + slide.annotationCount;
      if (rangeStart < nextOffset && rangeEnd > offset) {
        const localStart = Math.max(0, rangeStart - offset);
        const localEnd = Math.min(slide.annotationCount, rangeEnd - offset);
        requests.push(
          this.source
            .getAnnotations(slideIndex, signal)
            .then((items) => items.slice(localStart, localEnd)),
        );
      }
      offset = nextOffset;
    });
    return (await Promise.all(requests)).flat();
  }

  async findIndexById(id: string, signal?: AbortSignal) {
    let offset = 0;
    const { slides } = this.source.getSnapshot();
    for (let index = 0; index < slides.length; index += 1) {
      if (signal?.aborted) return -1;
      const count = slides[index].annotationCount;
      if (!count) continue;
      const annotations = await this.source.getAnnotations(index, signal);
      const localIndex = annotations.findIndex((item) => item.id === id);
      if (localIndex >= 0) return offset + localIndex;
      offset += count;
    }
    return -1;
  }

  /** 解除对演示文稿 Source 的订阅，不释放 Source 本身。 */
  dispose() {
    this.unsubscribe();
    this.listeners.clear();
  }

  private createSignature() {
    return this.source
      .getSnapshot()
      .slides.map((slide) => `${slide.id}:${slide.annotationCount}`)
      .join('|');
  }

  private createSnapshot(): OfficeAnnotationSourceSnapshot {
    return {
      revision: this.revision,
      count: this.source
        .getSnapshot()
        .slides.reduce((total, slide) => total + slide.annotationCount, 0),
      revisionCount: 0,
      noteCount: 0,
      supportsRevisionModes: false,
    };
  }

  private sync() {
    const nextSignature = this.createSignature();
    if (nextSignature === this.signature) return;
    this.signature = nextSignature;
    this.revision += 1;
    this.snapshot = this.createSnapshot();
    this.listeners.forEach((listener) => listener());
  }
}
