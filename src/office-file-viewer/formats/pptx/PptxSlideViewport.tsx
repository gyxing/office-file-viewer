// PptxSlideViewport 渲染当前幻灯片的滚动视口，并在翻页或缩放时复位滚动位置。
import React, { memo, useEffect, useRef } from 'react';
import type { SlideModel } from '../../services/pptx/types';
import type { OfficeFileViewerPresentationMediaOptions } from '../../services/presentation/mediaTypes';
import type {
  OfficeFileViewerPresentationTransitions,
  PresentationNavigationIntent,
} from '../../services/presentation/transitionTypes';
import { PresentationSlideState } from './PresentationSlideState';
import { PresentationTransitionLayer } from './PresentationTransitionLayer';

/** PPTX幻灯片视口组件属性。 */
type PptxSlideViewportProps = {
  /** 当前处理或展示的幻灯片。 */
  slide?: SlideModel;
  /** 当前选中项在所属集合中的索引。 */
  activeIndex: number;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 未加载出 Slide 时用于保持视口比例的文稿宽度。 */
  width?: number;
  /** 未加载出 Slide 时用于保持视口比例的文稿高度。 */
  height?: number;
  /** 当前页是否仍在按需读取。 */
  loading?: boolean;
  /** 当前页按需读取失败时的原始错误。 */
  error?: Error;
  /** 重试当前页读取。 */
  onRetry?: () => void;
  /** 演示文稿媒体读取配置。 */
  mediaOptions?: false | OfficeFileViewerPresentationMediaOptions;
  /** 是否按源文件播放页级切换。 */
  transitions: OfficeFileViewerPresentationTransitions;
  /** 最近一次工具栏翻页产生的切换意图。 */
  transitionIntent?: PresentationNavigationIntent;
};

/** 在可缩放视口中居中展示当前幻灯片。 */
function PptxSlideViewportComponent({
  slide,
  activeIndex,
  zoom,
  width = 960,
  height = 540,
  loading = false,
  error,
  onRetry = () => undefined,
  mediaOptions,
  transitions,
  transitionIntent,
}: PptxSlideViewportProps) {
  const viewportRef = useRef<HTMLElement | null>(null);

  useEffect(() => {
    viewportRef.current?.scrollTo({ left: 0, top: 0 });
  }, [activeIndex, zoom]);

  return (
    <section ref={viewportRef} className="office-file-pptx-viewer__viewport">
      <div className="office-file-pptx-viewer__slide-wrap">
        {error ? (
          <PresentationSlideState
            width={width}
            height={height}
            error={error}
            onRetry={onRetry}
          />
        ) : (
          <PresentationTransitionLayer
            slide={slide}
            activeIndex={activeIndex}
            zoom={zoom}
            intent={transitionIntent}
            transitions={transitions}
            mediaOptions={mediaOptions}
            fallback={
              loading ? (
                <PresentationSlideState
                  width={width}
                  height={height}
                  onRetry={onRetry}
                />
              ) : null
            }
          />
        )}
      </div>
    </section>
  );
}

export const PptxSlideViewport = memo(PptxSlideViewportComponent);
