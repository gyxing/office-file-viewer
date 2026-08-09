import React, {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  type KeyboardEvent,
} from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type {
  PresentationSlideDescriptor,
  PresentationSourceSnapshot,
} from '../../services/presentation/types';
import { usePresentationThumbnailWindow } from './usePresentationThumbnailWindow';

/** 演示文稿缩略图虚拟列表组件属性。 */
type VirtualPptxThumbnailListProps = {
  /** 当前数据源的只读快照。 */
  snapshot: PresentationSourceSnapshot;
  /** 当前激活项的零基索引。 */
  activeIndex: number;
  /** 用户选择幻灯片时触发的回调。 */
  onSelectSlide: (index: number) => void;
  /** 用于渲染指定缩略图的函数。 */
  renderThumbnail: (
    descriptor: PresentationSlideDescriptor,
    index: number,
  ) => React.ReactNode;
};

/** 根据键盘操作计算缩略图导航目标，未处理的按键不改变当前项。 */
export function getPresentationThumbnailNavigationTarget(
  key: string,
  currentIndex: number,
  itemCount: number,
) {
  if (itemCount <= 0) return undefined;
  if (key === 'ArrowUp') return Math.max(0, currentIndex - 1);
  if (key === 'ArrowDown') return Math.min(itemCount - 1, currentIndex + 1);
  if (key === 'Home') return 0;
  if (key === 'End') return itemCount - 1;
  return undefined;
}

/** 仅挂载侧栏可见窗口及上下预取范围内的缩略图。 */
export function VirtualPptxThumbnailList({
  snapshot,
  activeIndex,
  onSelectSlide,
  renderThumbnail,
}: VirtualPptxThumbnailListProps) {
  const messages = useOfficeFileViewerMessages();
  const viewportRef = useRef<HTMLDivElement>(null);
  const activeButtonRef = useRef<HTMLButtonElement | null>(null);
  const buttonRefs = useRef<Record<number, HTMLButtonElement | null>>({});
  const pendingFocusIndexRef = useRef<number>();
  const scrollGenerationRef = useRef(0);
  const itemHeight = useMemo(() => {
    const canvasWidth = 244;
    const ratio = snapshot.width / Math.max(1, snapshot.height);
    return Math.max(150, canvasWidth / ratio + 49);
  }, [snapshot.height, snapshot.width]);
  const range = usePresentationThumbnailWindow(
    viewportRef,
    snapshot.slideCount,
    itemHeight,
  );
  const handleThumbnailKeyDown = useCallback(
    (event: KeyboardEvent<HTMLButtonElement>, index: number) => {
      const targetIndex = getPresentationThumbnailNavigationTarget(
        event.key,
        index,
        snapshot.slideCount,
      );
      if (targetIndex === undefined) return;
      event.preventDefault();
      pendingFocusIndexRef.current = targetIndex;
      onSelectSlide(targetIndex);
      window.requestAnimationFrame(() => {
        const target = buttonRefs.current[targetIndex];
        if (target) {
          target.focus();
          pendingFocusIndexRef.current = undefined;
        }
      });
    },
    [onSelectSlide, snapshot.slideCount],
  );

  useEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return undefined;
    const generation = ++scrollGenerationRef.current;
    const frame = window.requestAnimationFrame(() => {
      if (generation !== scrollGenerationRef.current) return;
      const top = activeIndex * itemHeight;
      const bottom = top + itemHeight;
      if (
        top < viewport.scrollTop ||
        bottom > viewport.scrollTop + viewport.clientHeight
      ) {
        viewport.scrollTo({
          top: Math.max(0, top - viewport.clientHeight / 2 + itemHeight / 2),
        });
      }
      window.requestAnimationFrame(() => {
        if (generation === scrollGenerationRef.current) {
          activeButtonRef.current?.scrollIntoView({ block: 'nearest' });
        }
      });
    });
    return () => {
      scrollGenerationRef.current += 1;
      window.cancelAnimationFrame(frame);
    };
  }, [activeIndex, itemHeight]);

  useEffect(() => {
    const targetIndex = pendingFocusIndexRef.current;
    if (
      targetIndex === undefined ||
      targetIndex < range.start ||
      targetIndex >= range.end
    ) {
      return undefined;
    }
    const frame = window.requestAnimationFrame(() => {
      const target = buttonRefs.current[targetIndex];
      if (target) {
        target.focus();
        pendingFocusIndexRef.current = undefined;
      }
    });
    return () => window.cancelAnimationFrame(frame);
  }, [activeIndex, range.end, range.start]);

  return (
    <div
      ref={viewportRef}
      className="office-file-pptx-viewer__thumbnail-list"
      data-thumbnail-mode="virtual"
      data-mounted-thumbnail-count={range.end - range.start}
    >
      <div style={{ height: range.topSpacer }} aria-hidden="true" />
      {snapshot.slides
        .slice(range.start, range.end)
        .map((descriptor, offset) => {
          const index = range.start + offset;
          return (
            <button
              ref={(node) => {
                buttonRefs.current[index] = node;
                if (index === activeIndex) activeButtonRef.current = node;
              }}
              key={descriptor.id}
              type="button"
              aria-label={
                descriptor.status === 'error'
                  ? `${messages.presentation.slide(descriptor.index)}，${
                      messages.lazyContent.retry
                    }`
                  : messages.presentation.slide(descriptor.index)
              }
              aria-current={index === activeIndex ? 'page' : undefined}
              aria-posinset={index + 1}
              aria-setsize={snapshot.slideCount}
              tabIndex={index === activeIndex ? 0 : -1}
              className="office-file-pptx-viewer__thumbnail-button"
              style={{ minHeight: itemHeight - 12 }}
              onClick={() => onSelectSlide(index)}
              onKeyDown={(event) => handleThumbnailKeyDown(event, index)}
            >
              {renderThumbnail(descriptor, index)}
            </button>
          );
        })}
      <div style={{ height: range.bottomSpacer }} aria-hidden="true" />
    </div>
  );
}
