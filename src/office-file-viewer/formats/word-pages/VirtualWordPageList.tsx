import type { MutableRefObject, ReactNode, RefObject } from 'react';
import React, {
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
} from 'react';
import type { WordPageSource } from '../../services/word/WordPageSource';
import './index.less';
import { PageHeightIndex } from './PageHeightIndex';
import type {
  WordPageLoadState,
  WordPageNavigationController,
  WordPageWindowRange,
} from './types';
import { useWordPageWindow } from './useWordPageWindow';
import { WordPagePlaceholder } from './WordPagePlaceholder';

type VirtualWordPageListProps<TPage> = {
  source: WordPageSource<TPage>;
  scrollerRef: RefObject<HTMLElement>;
  layoutRevision: string;
  zoom: number;
  pageGap?: number;
  navigationControllerRef?: MutableRefObject<
    WordPageNavigationController | undefined
  >;
  renderPage(page: TPage, pageIndex: number): ReactNode;
};

type PendingAnchor = {
  pageIndex: number;
  viewportOffset: number;
};

function isAbortError(error: unknown) {
  return error instanceof Error && error.name === 'AbortError';
}

function throwIfAborted(signal?: AbortSignal) {
  if (!signal?.aborted) return;
  const error = new Error('操作已取消');
  error.name = 'AbortError';
  throw error;
}

function nextAnimationFrame(signal?: AbortSignal) {
  return new Promise<void>((resolve, reject) => {
    throwIfAborted(signal);
    const frame = requestAnimationFrame(() => {
      signal?.removeEventListener('abort', handleAbort);
      resolve();
    });
    function handleAbort() {
      cancelAnimationFrame(frame);
      const error = new Error('操作已取消');
      error.name = 'AbortError';
      reject(error);
    }
    signal?.addEventListener('abort', handleAbort, { once: true });
  });
}

/** 只挂载视口附近 Word 页面，并用两个 spacer 保持完整滚动高度。 */
export function VirtualWordPageList<TPage>({
  source,
  scrollerRef,
  layoutRevision,
  zoom,
  pageGap = 24,
  navigationControllerRef,
  renderPage,
}: VirtualWordPageListProps<TPage>) {
  const snapshot = useSyncExternalStore(
    source.subscribe,
    source.getSnapshot,
    source.getSnapshot,
  );
  const estimatedHeights = useMemo(
    () =>
      snapshot.pages.map(
        (meta, index) =>
          meta.estimatedContentHeight * (zoom / 100) +
          (index === snapshot.pages.length - 1 ? 0 : pageGap),
      ),
    [pageGap, snapshot.pages, zoom],
  );
  const heightIndex = useMemo(
    () => new PageHeightIndex(estimatedHeights),
    [estimatedHeights],
  );
  const [heightRevision, setHeightRevision] = useState(0);
  const [retryRevision, setRetryRevision] = useState(0);
  const [forcedPageIndex, setForcedPageIndex] = useState<number>();
  const [pageStates, setPageStates] = useState<
    Map<number, WordPageLoadState<TPage>>
  >(() => new Map());
  const listRef = useRef<HTMLDivElement>(null);
  const pendingAnchorRef = useRef<PendingAnchor>();
  const previousHeightIndexRef = useRef(heightIndex);
  const windowRange = useWordPageWindow(
    scrollerRef,
    heightIndex,
    `${layoutRevision}:${snapshot.revision}`,
    heightRevision,
  );
  const range = useMemo<WordPageWindowRange>(() => {
    if (
      forcedPageIndex === undefined ||
      (forcedPageIndex >= windowRange.start &&
        forcedPageIndex < windowRange.end)
    ) {
      return windowRange;
    }
    const pageCount = heightIndex.length;
    let start = Math.max(0, forcedPageIndex - 1);
    let end = Math.min(pageCount, start + 3);
    start = Math.max(0, end - 3);
    return {
      start,
      end,
      topSpacerHeight: heightIndex.prefix(start),
      bottomSpacerHeight: heightIndex.total() - heightIndex.prefix(end),
    };
  }, [forcedPageIndex, heightIndex, windowRange]);
  const rangeRef = useRef(range);
  rangeRef.current = range;

  const navigationController = useMemo<WordPageNavigationController>(
    () => ({
      scrollToPage(index, offset = 0) {
        const scroller = scrollerRef.current;
        if (!scroller || index < 0 || index >= heightIndex.length) return;
        const targetOffset = heightIndex.prefix(index) + offset;
        const usesInternalScroll =
          scroller.scrollHeight > scroller.clientHeight + 1;
        if (usesInternalScroll) {
          scroller.scrollTo({ top: targetOffset, behavior: 'auto' });
        } else {
          const currentOffset = Math.max(
            0,
            -scroller.getBoundingClientRect().top,
          );
          window.scrollBy({
            top: targetOffset - currentOffset,
            behavior: 'auto',
          });
        }
      },
      async ensurePageMounted(index, signal) {
        throwIfAborted(signal);
        await source.ensureRange(index, index + 1, signal);
        setForcedPageIndex(index);
        for (let attempt = 0; attempt < 120; attempt += 1) {
          await nextAnimationFrame(signal);
          const element = listRef.current?.querySelector<HTMLElement>(
            `[data-office-word-page-index="${index}"]`,
          );
          if (
            element &&
            !element.querySelector('.office-file-word-pages__placeholder')
          ) {
            requestAnimationFrame(() =>
              setForcedPageIndex((current) =>
                current === index ? undefined : current,
              ),
            );
            return element;
          }
        }
        throw new Error('目标页面未能及时挂载');
      },
      getMountedRange: () => ({
        start: rangeRef.current.start,
        end: rangeRef.current.end,
      }),
    }),
    [heightIndex, scrollerRef, source],
  );

  useLayoutEffect(() => {
    if (!navigationControllerRef) return;
    navigationControllerRef.current = navigationController;
    return () => {
      if (navigationControllerRef.current === navigationController) {
        navigationControllerRef.current = undefined;
      }
    };
  }, [navigationController, navigationControllerRef]);

  useLayoutEffect(() => {
    const previousIndex = previousHeightIndexRef.current;
    previousHeightIndexRef.current = heightIndex;
    const scroller = scrollerRef.current;
    if (previousIndex === heightIndex || !scroller) return;
    const usesInternalScroll =
      scroller.scrollHeight > scroller.clientHeight + 1;
    const currentOffset = usesInternalScroll
      ? scroller.scrollTop
      : Math.max(0, -scroller.getBoundingClientRect().top);
    const pageIndex = previousIndex.findIndexAtOffset(currentOffset);
    if (pageIndex < 0 || pageIndex >= heightIndex.length) return;
    const previousHeight =
      previousIndex.prefix(pageIndex + 1) - previousIndex.prefix(pageIndex);
    const relativeOffset =
      previousHeight > 0
        ? (currentOffset - previousIndex.prefix(pageIndex)) / previousHeight
        : 0;
    const nextHeight =
      heightIndex.prefix(pageIndex + 1) - heightIndex.prefix(pageIndex);
    const nextOffset =
      heightIndex.prefix(pageIndex) +
      Math.max(0, Math.min(1, relativeOffset)) * nextHeight;
    const correction = nextOffset - currentOffset;
    if (Math.abs(correction) <= 0.5) return;
    if (usesInternalScroll) scroller.scrollTop = nextOffset;
    else window.scrollBy(0, correction);
  }, [heightIndex, scrollerRef]);

  useEffect(() => {
    const controller = new AbortController();
    const releaseRange = source.retainRange(range.start, range.end);
    const metas = snapshot.pages.slice(range.start, range.end);
    setPageStates((current) => {
      const next = new Map<number, WordPageLoadState<TPage>>();
      metas.forEach((meta) => {
        const existing = current.get(meta.index);
        if (
          existing &&
          existing.status !== 'loading' &&
          existing.revision === meta.revision
        ) {
          next.set(meta.index, existing);
        } else {
          next.set(meta.index, { status: 'loading' });
        }
      });
      return next;
    });

    void source
      .ensureRange(range.start, range.end, controller.signal)
      .then(async () => {
        const results = await Promise.all(
          metas.map(async (meta) => {
            if (meta.status === 'error') {
              return {
                index: meta.index,
                state: {
                  status: 'error',
                  error: meta.errorMessage,
                  revision: meta.revision,
                } satisfies WordPageLoadState<TPage>,
              };
            }
            try {
              const page = await source.getPage(meta.index, controller.signal);
              if (page === undefined) throw new Error('页面内容尚未就绪');
              return {
                index: meta.index,
                state: {
                  status: 'ready',
                  page,
                  revision: meta.revision,
                } satisfies WordPageLoadState<TPage>,
              };
            } catch (error) {
              if (isAbortError(error)) throw error;
              return {
                index: meta.index,
                state: {
                  status: 'error',
                  error,
                  revision: meta.revision,
                } satisfies WordPageLoadState<TPage>,
              };
            }
          }),
        );
        if (controller.signal.aborted) return;
        setPageStates(
          new Map(results.map((result) => [result.index, result.state])),
        );
      })
      .catch((error) => {
        if (isAbortError(error) || controller.signal.aborted) return;
        setPageStates(
          new Map(
            metas.map((meta) => [
              meta.index,
              {
                status: 'error',
                error,
                revision: meta.revision,
              } satisfies WordPageLoadState<TPage>,
            ]),
          ),
        );
      });

    return () => {
      controller.abort();
      releaseRange();
    };
  }, [range.end, range.start, retryRevision, snapshot.pages, source]);

  useLayoutEffect(() => {
    const anchor = pendingAnchorRef.current;
    const scroller = scrollerRef.current;
    const list = listRef.current;
    if (!anchor || !scroller || !list) return;
    pendingAnchorRef.current = undefined;
    const element = list.querySelector<HTMLElement>(
      `[data-office-word-page-index="${anchor.pageIndex}"]`,
    );
    if (!element) return;
    const usesInternalScroll =
      scroller.scrollHeight > scroller.clientHeight + 1;
    const viewportTop = usesInternalScroll
      ? scroller.getBoundingClientRect().top
      : 0;
    const nextOffset = element.getBoundingClientRect().top - viewportTop;
    const correction = nextOffset - anchor.viewportOffset;
    if (Math.abs(correction) <= 0.5) return;
    if (usesInternalScroll) scroller.scrollTop += correction;
    else window.scrollBy(0, correction);
  }, [heightRevision, scrollerRef]);

  useEffect(() => {
    const scroller = scrollerRef.current;
    const list = listRef.current;
    if (!scroller || !list || typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver((entries) => {
      const scrollerRect = scroller.getBoundingClientRect();
      const usesInternalScroll =
        scroller.scrollHeight > scroller.clientHeight + 1;
      const viewportTop = usesInternalScroll ? scrollerRect.top : 0;
      const mountedItems = Array.from(
        list.querySelectorAll<HTMLElement>('[data-office-word-page-index]'),
      );
      const anchor = mountedItems.find(
        (element) => element.getBoundingClientRect().bottom > viewportTop,
      );
      let changed = false;
      entries.forEach((entry) => {
        const element = entry.target as HTMLElement;
        const pageIndex = Number(element.dataset.officeWordPageIndex);
        if (!Number.isInteger(pageIndex)) return;
        const measuredHeight = element.getBoundingClientRect().height;
        if (
          Math.abs(
            heightIndex.prefix(pageIndex + 1) -
              heightIndex.prefix(pageIndex) -
              measuredHeight,
          ) <= 0.5
        ) {
          return;
        }
        heightIndex.replace(pageIndex, measuredHeight);
        changed = true;
      });
      if (!changed) return;
      if (anchor) {
        pendingAnchorRef.current = {
          pageIndex: Number(anchor.dataset.officeWordPageIndex),
          viewportOffset: anchor.getBoundingClientRect().top - viewportTop,
        };
      }
      setHeightRevision((current) => current + 1);
    });
    list
      .querySelectorAll<HTMLElement>('[data-office-word-page-index]')
      .forEach((element) => observer.observe(element));
    return () => observer.disconnect();
  }, [heightIndex, pageStates, range.end, range.start, scrollerRef]);

  const mountedMetas = snapshot.pages.slice(range.start, range.end);
  return (
    <div
      ref={listRef}
      className="office-file-word-pages"
      data-page-count={snapshot.pageCount ?? snapshot.pages.length}
      data-mounted-page-count={mountedMetas.length}
    >
      <div
        className="office-file-word-pages__spacer"
        style={{ height: range.topSpacerHeight }}
        aria-hidden="true"
      />
      {mountedMetas.map((meta) => {
        const state = pageStates.get(meta.index);
        const estimatedHeight =
          estimatedHeights[meta.index] ?? meta.estimatedContentHeight;
        const trailingGap =
          meta.index === snapshot.pages.length - 1 ? 0 : pageGap;
        return (
          <div
            key={meta.id}
            className="office-file-word-pages__item"
            data-office-word-page-index={meta.index}
            style={{ paddingBottom: trailingGap }}
          >
            {state?.status === 'ready' ? (
              renderPage(state.page, meta.index)
            ) : (
              <WordPagePlaceholder
                status={state?.status === 'error' ? 'error' : 'loading'}
                minHeight={Math.max(120, estimatedHeight - trailingGap)}
                onRetry={
                  state?.status === 'error'
                    ? () => {
                        source.retry(meta.index);
                        setRetryRevision((current) => current + 1);
                      }
                    : undefined
                }
              />
            )}
          </div>
        );
      })}
      <div
        className="office-file-word-pages__spacer"
        style={{ height: range.bottomSpacerHeight }}
        aria-hidden="true"
      />
    </div>
  );
}
