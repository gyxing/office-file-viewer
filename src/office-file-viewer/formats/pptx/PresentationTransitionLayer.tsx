import type { AnimationEvent, CSSProperties, ReactNode } from 'react';
import React, { useEffect, useMemo, useRef, useState } from 'react';
import type { OfficeFileViewerPresentationMediaOptions } from '../../services/presentation/mediaTypes';
import type {
  OfficeFileViewerPresentationTransitions,
  PresentationNavigationIntent,
} from '../../services/presentation/transitionTypes';
import type { SlideModel } from '../../services/presentation/types';
import { PptxSlide } from './PptxSlide';

type ActiveTransition = Readonly<{
  key: number;
  outgoing: SlideModel;
  className: string;
  durationMs: number;
}>;

function prefersReducedMotion() {
  return (
    typeof window !== 'undefined' &&
    typeof window.matchMedia === 'function' &&
    window.matchMedia('(prefers-reduced-motion: reduce)').matches
  );
}

function transitionClass(
  slide: SlideModel,
  intent: PresentationNavigationIntent,
) {
  const transition = slide.transition!;
  const direction =
    transition.direction ?? (intent.direction === 'next' ? 'left' : 'right');
  return `office-file-presentation-transition--${transition.type} office-file-presentation-transition--${direction}`;
}

/** 页级切换层属性。 */
type PresentationTransitionLayerProps = {
  /** 当前已加载幻灯片；按需加载期间可以暂时为空。 */
  slide?: SlideModel;
  /** 当前目标幻灯片零基索引。 */
  activeIndex: number;
  /** 当前缩放比例。 */
  zoom: number;
  /** 仅工具栏上一页/下一页产生的切换意图。 */
  intent?: PresentationNavigationIntent;
  /** 是否按源文件播放页级切换。 */
  transitions: OfficeFileViewerPresentationTransitions;
  /** 演示文稿媒体读取配置。 */
  mediaOptions?: false | OfficeFileViewerPresentationMediaOptions;
  /** 尚无任何可显示页面时的占位内容。 */
  fallback?: ReactNode;
};

/** 在快速翻页时可取消旧动画，并在完成后只保留当前幻灯片。 */
export function PresentationTransitionLayer({
  slide,
  activeIndex,
  zoom,
  intent,
  transitions,
  mediaOptions,
  fallback,
}: PresentationTransitionLayerProps) {
  const displayedRef = useRef<SlideModel | undefined>(slide);
  const consumedTokenRef = useRef(0);
  const timerRef = useRef<number>();
  const [displayed, setDisplayed] = useState<SlideModel | undefined>(slide);
  const [activeTransition, setActiveTransition] = useState<ActiveTransition>();

  useEffect(() => {
    if (!slide) return;
    const previous = displayedRef.current;
    displayedRef.current = slide;
    setDisplayed(slide);
    if (!previous || previous.id === slide.id) return;
    if (timerRef.current !== undefined) window.clearTimeout(timerRef.current);
    const shouldAnimate =
      transitions === 'source' &&
      Boolean(slide.transition) &&
      Boolean(intent) &&
      intent!.targetIndex === activeIndex &&
      intent!.token > consumedTokenRef.current &&
      !prefersReducedMotion();
    if (intent?.targetIndex === activeIndex) {
      consumedTokenRef.current = Math.max(
        consumedTokenRef.current,
        intent.token,
      );
    }
    if (!shouldAnimate) {
      setActiveTransition(undefined);
      return;
    }
    const next: ActiveTransition = {
      key: intent!.token,
      outgoing: previous,
      className: transitionClass(slide, intent!),
      durationMs: slide.transition!.durationMs,
    };
    setActiveTransition(next);
    timerRef.current = window.setTimeout(
      () =>
        setActiveTransition((current) =>
          current?.key === next.key ? undefined : current,
        ),
      next.durationMs + 80,
    );
  }, [activeIndex, intent, slide, transitions]);

  useEffect(
    () => () => {
      if (timerRef.current !== undefined) window.clearTimeout(timerRef.current);
    },
    [],
  );

  const style = useMemo<
    CSSProperties & {
      '--office-file-presentation-transition-duration': string;
    }
  >(
    () => ({
      width: (displayed?.width ?? 0) * (zoom / 100),
      height: (displayed?.height ?? 0) * (zoom / 100),
      '--office-file-presentation-transition-duration': `${
        activeTransition?.durationMs ?? 0
      }ms`,
    }),
    [activeTransition?.durationMs, displayed?.height, displayed?.width, zoom],
  );
  if (!displayed) return <>{fallback}</>;

  return (
    <div
      className={[
        'office-file-presentation-transition',
        activeTransition?.className,
      ]
        .filter(Boolean)
        .join(' ')}
      style={style}
      data-office-fit-target="true"
      data-office-fit-base-width={displayed.width}
      data-office-fit-base-height={displayed.height}
      data-transition-active={activeTransition ? 'true' : 'false'}
      data-transition-type={displayed.transition?.type}
      data-transition-intent={
        intent?.targetIndex === activeIndex ? intent.token : undefined
      }
    >
      {activeTransition ? (
        <div className="office-file-presentation-transition__slide office-file-presentation-transition__slide--outgoing">
          <PptxSlide
            key={`outgoing-${activeTransition.key}`}
            slide={activeTransition.outgoing}
            zoom={zoom}
            renderKey={`transition-out-${activeTransition.outgoing.id}`}
            interactive={false}
            mediaOptions={false}
          />
        </div>
      ) : null}
      <div
        key={`incoming-${activeTransition?.key ?? displayed.id}`}
        className="office-file-presentation-transition__slide office-file-presentation-transition__slide--incoming"
        onAnimationEnd={(event: AnimationEvent<HTMLDivElement>) => {
          // 媒体控件或页内元素的动画会冒泡，只有切换包装层自身结束才清理旧页。
          if (event.currentTarget === event.target) {
            setActiveTransition(undefined);
          }
        }}
      >
        <PptxSlide
          slide={displayed}
          zoom={zoom}
          renderKey={`slide-${displayed.id}`}
          searchSlideIndex={activeIndex}
          mediaOptions={mediaOptions}
        />
      </div>
    </div>
  );
}
