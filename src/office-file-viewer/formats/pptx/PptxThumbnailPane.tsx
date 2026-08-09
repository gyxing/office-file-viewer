// PptxThumbnailPane 渲染幻灯片缩略图列表，并负责切换当前页。
import React, {
  memo,
  useCallback,
  useEffect,
  useRef,
  useState,
  type KeyboardEvent,
} from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type {
  PresentationSlideDescriptor,
  PresentationSource,
  PresentationSourceSnapshot,
  SlideModel,
} from '../../services/presentation/types';
import { PptxThumbnail } from './PptxThumbnail';
import { PptxThumbnailPlaceholder } from './PptxThumbnailPlaceholder';
import {
  getPresentationThumbnailNavigationTarget,
  VirtualPptxThumbnailList,
} from './VirtualPptxThumbnailList';

/** PPTX缩略图面板组件属性。 */
type PptxThumbnailPaneProps = {
  /** 当前预览使用的按需加载数据源。 */
  source: PresentationSource;
  /** 当前数据源的只读快照。 */
  snapshot: PresentationSourceSnapshot;
  /** 当前激活项的零基索引。 */
  activeIndex: number;
  /** 用户选择幻灯片时触发的回调。 */
  onSelectSlide: (index: number) => void;
};

/** 单张演示文稿缩略图内容组件属性。 */
type ThumbnailContentProps = {
  /** 当前预览使用的按需加载数据源。 */
  source: PresentationSource;
  /** 当前项目的轻量描述信息。 */
  descriptor: PresentationSlideDescriptor;
  /** 在所属集合中的零基索引。 */
  index: number;
  /** 当前项目是否处于激活状态。 */
  active: boolean;
  /** 是否跳过可见性等待并立即读取缩略图。 */
  shouldLoadImmediately: boolean;
};

/** 按描述符状态和可见性读取一页缩略图模型。 */
function ThumbnailContent({
  source,
  descriptor,
  index,
  active,
  shouldLoadImmediately,
}: ThumbnailContentProps) {
  const observerTargetRef = useRef<HTMLDivElement>(null);
  const [intersecting, setIntersecting] = useState(shouldLoadImmediately);
  const [slide, setSlide] = useState<SlideModel>();

  useEffect(() => {
    if (shouldLoadImmediately) {
      setIntersecting(true);
      return undefined;
    }
    const target = observerTargetRef.current;
    if (!target || typeof IntersectionObserver === 'undefined') {
      setIntersecting(true);
      return undefined;
    }
    const observer = new IntersectionObserver(
      (entries) => {
        if (entries.some((entry) => entry.isIntersecting)) {
          setIntersecting(true);
          observer.disconnect();
        }
      },
      { rootMargin: '240px 0px' },
    );
    observer.observe(target);
    return () => observer.disconnect();
  }, [shouldLoadImmediately]);

  useEffect(() => {
    if (!intersecting) return undefined;
    const controller = new AbortController();
    void source.getSlide(index, controller.signal).then(
      (nextSlide) => {
        if (!controller.signal.aborted) setSlide(nextSlide);
      },
      () => {
        // Source 会把失败写回 descriptor；占位组件统一呈现错误和重试入口。
      },
    );
    return () => controller.abort();
  }, [descriptor.revision, index, intersecting, source]);

  return (
    <div ref={observerTargetRef}>
      {slide ? (
        <PptxThumbnail slide={slide} active={active} />
      ) : (
        <PptxThumbnailPlaceholder
          descriptor={descriptor}
          aspectRatio={source.getSnapshot().width / source.getSnapshot().height}
          active={active}
        />
      )}
    </div>
  );
}

/** 渲染幻灯片缩略图导航并切换当前页。 */
function PptxThumbnailPaneComponent({
  source,
  snapshot,
  activeIndex,
  onSelectSlide,
}: PptxThumbnailPaneProps) {
  const messages = useOfficeFileViewerMessages();
  const buttonRefs = useRef<Record<number, HTMLButtonElement | null>>({});
  const handleSelect = useCallback(
    (index: number, descriptor: PresentationSlideDescriptor) => {
      if (descriptor.status === 'error') {
        source.retry(index);
        return;
      }
      onSelectSlide(index);
    },
    [onSelectSlide, source],
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
      handleSelect(targetIndex, snapshot.slides[targetIndex]);
      window.requestAnimationFrame(() =>
        buttonRefs.current[targetIndex]?.focus(),
      );
    },
    [handleSelect, snapshot.slideCount, snapshot.slides],
  );
  const renderThumbnail = useCallback(
    (descriptor: PresentationSlideDescriptor, index: number) => (
      <ThumbnailContent
        source={source}
        descriptor={descriptor}
        index={index}
        active={index === activeIndex}
        shouldLoadImmediately
      />
    ),
    [activeIndex, source],
  );

  return (
    <aside
      className="office-file-pptx-viewer__sidebar"
      aria-label={messages.presentation.thumbnailsRegion}
    >
      <div className="office-file-pptx-viewer__sidebar-header">
        <div className="office-file-pptx-viewer__slide-count">
          {messages.presentation.slideCount(snapshot.slideCount)}
        </div>
      </div>
      {snapshot.performance.thumbnailMode === 'virtual' ? (
        <VirtualPptxThumbnailList
          snapshot={snapshot}
          activeIndex={activeIndex}
          onSelectSlide={(index) => handleSelect(index, snapshot.slides[index])}
          renderThumbnail={renderThumbnail}
        />
      ) : (
        <div
          className="office-file-pptx-viewer__thumbnail-list"
          data-thumbnail-mode="normal"
          data-mounted-thumbnail-count={snapshot.slideCount}
        >
          {snapshot.slides.map((descriptor, index) => (
            <button
              ref={(node) => {
                buttonRefs.current[index] = node;
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
              onClick={() => handleSelect(index, descriptor)}
              onKeyDown={(event) => handleThumbnailKeyDown(event, index)}
            >
              <ThumbnailContent
                source={source}
                descriptor={descriptor}
                index={index}
                active={index === activeIndex}
                shouldLoadImmediately={
                  snapshot.performance.slideMode === 'materialized' ||
                  Math.abs(index - activeIndex) <= 2
                }
              />
            </button>
          ))}
        </div>
      )}
    </aside>
  );
}

export const PptxThumbnailPane = memo(PptxThumbnailPaneComponent);
