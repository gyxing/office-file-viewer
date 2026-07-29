import { useCallback, useEffect, useState, useSyncExternalStore } from 'react';
import { PRESENTATION_PRELOAD_RADIUS } from '../../services/presentation/presentationPerformance';
import type {
  PresentationSource,
  PresentationSourceSnapshot,
  SlideModel,
  SpeakerNotesModel,
} from '../../services/presentation/types';

const EMPTY_PRESENTATION_SNAPSHOT: PresentationSourceSnapshot = {
  revision: 0,
  width: 960,
  height: 540,
  theme: { colorScheme: {}, fontScheme: {} },
  slideCount: 0,
  slides: [],
  performance: {
    thumbnailMode: 'normal',
    slideMode: 'materialized',
    totalElementWeight: 0,
  },
};

type PresentationContentState = {
  slide?: SlideModel;
  notes?: SpeakerNotesModel;
  loading: boolean;
  error?: Error;
};

/** 订阅 PresentationSource，并只加载当前页、邻近预取页和按需备注。 */
export function usePresentationSource(
  source: PresentationSource | undefined,
  activeIndex: number,
  showSpeakerNotes: boolean,
) {
  const subscribe = useCallback(
    (listener: () => void) =>
      source ? source.subscribe(listener) : () => undefined,
    [source],
  );
  const getSnapshot = useCallback(
    () => source?.getSnapshot() ?? EMPTY_PRESENTATION_SNAPSHOT,
    [source],
  );
  const snapshot = useSyncExternalStore(
    subscribe,
    getSnapshot,
    () => EMPTY_PRESENTATION_SNAPSHOT,
  );
  const [retryRevision, setRetryRevision] = useState(0);
  const [state, setState] = useState<PresentationContentState>({
    loading: Boolean(source && snapshot.slideCount),
  });

  useEffect(() => {
    if (!source || activeIndex < 0 || activeIndex >= snapshot.slideCount) {
      setState({ loading: false });
      return undefined;
    }

    const controller = new AbortController();
    const start = Math.max(0, activeIndex - PRESENTATION_PRELOAD_RADIUS);
    const end = Math.min(
      snapshot.slideCount - 1,
      activeIndex + PRESENTATION_PRELOAD_RADIUS,
    );
    const release = source.retainRange(start, end);
    setState((current) => ({
      slide:
        current.slide?.index === snapshot.slides[activeIndex]?.index
          ? current.slide
          : undefined,
      loading: true,
    }));

    void source.ensureRange(start, end, controller.signal).catch(() => {
      // 邻近页预取失败不覆盖当前页自己的 loading/error 状态。
    });
    void source.getSlide(activeIndex, controller.signal).then(
      (slide) => {
        if (controller.signal.aborted) return;
        setState((current) => ({
          ...current,
          slide,
          loading: false,
          error: undefined,
        }));
      },
      (error) => {
        if (controller.signal.aborted) return;
        setState({
          loading: false,
          error: error instanceof Error ? error : new Error('幻灯片加载失败'),
        });
      },
    );

    return () => {
      controller.abort();
      release();
    };
  }, [activeIndex, retryRevision, snapshot.slideCount, source]);

  useEffect(() => {
    if (
      !source ||
      !showSpeakerNotes ||
      activeIndex < 0 ||
      activeIndex >= snapshot.slideCount
    ) {
      setState((current) =>
        current.notes === undefined
          ? current
          : { ...current, notes: undefined },
      );
      return undefined;
    }

    const controller = new AbortController();
    void source.getSpeakerNotes(activeIndex, controller.signal).then(
      (notes) => {
        if (!controller.signal.aborted) {
          setState((current) => ({ ...current, notes }));
        }
      },
      () => {
        // 备注失败不应让已成功加载的主幻灯片进入错误态。
      },
    );
    return () => controller.abort();
  }, [
    activeIndex,
    retryRevision,
    showSpeakerNotes,
    snapshot.slideCount,
    source,
  ]);

  const retry = useCallback(() => {
    source?.retry(activeIndex);
    setRetryRevision((value) => value + 1);
  }, [activeIndex, source]);

  return { snapshot, ...state, retry };
}
