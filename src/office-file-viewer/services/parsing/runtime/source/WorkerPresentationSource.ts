import type {
  PresentationSource,
  PresentationSourceSnapshot,
} from '../../../presentation/PresentationSource';
import type {
  PresentationAnnotation,
  SlideModel,
  SpeakerNotesModel,
} from '../../../presentation/types';
import type { OfficeSearchProvider } from '../../../search/types';
import type { WorkerPresentationSourceState } from '../../protocol/messages';
import { WorkerSourceClient } from './WorkerSourceClient';

/** 将 Worker 中长期持有的 PPTX Source 适配为现有幻灯片接口。 */
export class WorkerPresentationSource implements PresentationSource {
  readonly searchProvider: OfficeSearchProvider = {
    kind: 'presentation',
    search: (query, emit, signal) => this.client.search(query, emit, signal),
  };
  private readonly listeners = new Set<() => void>();
  private readonly unsubscribeUpdate: () => void;
  private readonly unsubscribeFailure: () => void;
  private snapshot: PresentationSourceSnapshot;
  private disposed = false;

  constructor(
    private readonly client: WorkerSourceClient,
    initial: WorkerPresentationSourceState,
  ) {
    this.snapshot = initial.snapshot;
    this.unsubscribeUpdate = client.subscribe((source) => {
      if (source.kind !== 'pptx') return;
      this.snapshot = source.snapshot;
      this.emitChange();
    });
    this.unsubscribeFailure = client.subscribeFailure((error) => {
      this.snapshot = {
        ...this.snapshot,
        revision: this.snapshot.revision + 1,
        slides: this.snapshot.slides.map((slide) =>
          slide.status === 'estimated'
            ? {
                ...slide,
                revision: slide.revision + 1,
                status: 'error',
                errorMessage: error.message,
              }
            : slide,
        ),
      };
      this.emitChange();
    });
  }

  getSnapshot = () => this.snapshot;

  subscribe = (listener: () => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  getSlide(index: number, signal?: AbortSignal) {
    return this.client.request<SlideModel>('get-slide', { index }, { signal });
  }

  getSpeakerNotes(index: number, signal?: AbortSignal) {
    return this.client.request<SpeakerNotesModel | undefined>(
      'get-notes',
      { index },
      { signal },
    );
  }

  getAnnotations(index: number, signal?: AbortSignal) {
    return this.client.request<readonly PresentationAnnotation[]>(
      'get-presentation-annotations',
      { index },
      { signal },
    );
  }

  ensureRange(start: number, end: number, signal?: AbortSignal) {
    return this.client.request<void>(
      'ensure-presentation-range',
      { start, end },
      { signal },
    );
  }

  retainRange(start: number, end: number) {
    return this.client.retain('retain-presentation-range', { start, end });
  }

  retry(index: number) {
    void this.client
      .request('retry-presentation-slide', { index })
      .catch(() => undefined);
  }

  async dispose() {
    if (this.disposed) return;
    this.disposed = true;
    this.unsubscribeUpdate();
    this.unsubscribeFailure();
    this.listeners.clear();
    await this.client.dispose();
  }

  private emitChange() {
    this.listeners.forEach((listener) => listener());
  }
}
