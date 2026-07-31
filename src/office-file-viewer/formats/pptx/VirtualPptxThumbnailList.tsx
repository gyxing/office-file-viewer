import React, { useEffect, useMemo, useRef } from 'react';
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

/** 仅挂载侧栏可见窗口及上下预取范围内的缩略图。 */
export function VirtualPptxThumbnailList({
  snapshot,
  activeIndex,
  onSelectSlide,
  renderThumbnail,
}: VirtualPptxThumbnailListProps) {
  const messages = useOfficeFileViewerMessages();
  const viewportRef = useRef<HTMLDivElement>(null);
  const activeButtonRef = useRef<HTMLButtonElement>(null);
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
              ref={index === activeIndex ? activeButtonRef : undefined}
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
              className="office-file-pptx-viewer__thumbnail-button"
              style={{ minHeight: itemHeight - 12 }}
              onClick={() => onSelectSlide(index)}
            >
              {renderThumbnail(descriptor, index)}
            </button>
          );
        })}
      <div style={{ height: range.bottomSpacer }} aria-hidden="true" />
    </div>
  );
}
