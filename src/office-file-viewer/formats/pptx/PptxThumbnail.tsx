// PptxThumbnail 复用单页幻灯片渲染能力，生成缩略图预览。
import type { CSSProperties } from 'react';
import React, { memo, useEffect, useMemo, useRef, useState } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type { SlideModel } from '../../services/pptx/types';
import { PptxSlide } from './PptxSlide';
import { colorWithOpacity } from './renderers/paint';

/** PPTX缩略图组件属性。 */
type PptxThumbnailProps = {
  /** 当前处理或展示的幻灯片。 */
  slide: SlideModel;
  /** 当前项目是否处于选中状态。 */
  active: boolean;
};

/** 渲染PPTX缩略图。 */
function PptxThumbnailComponent({ slide, active }: PptxThumbnailProps) {
  const messages = useOfficeFileViewerMessages();
  const canvasRef = useRef<HTMLDivElement>(null);
  const [thumbnailScale, setThumbnailScale] = useState(0.18);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) {
      return undefined;
    }

    const updateScale = () => {
      const nextScale = canvas.clientWidth / slide.width;
      setThumbnailScale((currentScale) =>
        Math.abs(currentScale - nextScale) < 0.0001 ? currentScale : nextScale,
      );
    };

    updateScale();
    if (typeof ResizeObserver === 'undefined') {
      return undefined;
    }

    const observer = new ResizeObserver(updateScale);
    observer.observe(canvas);
    return () => observer.disconnect();
  }, [slide.width]);

  const canvasStyle = useMemo<CSSProperties>(
    () => ({
      aspectRatio: `${slide.width / slide.height}`,
      background: colorWithOpacity(
        slide.background?.fill ?? '#f8fafc',
        slide.background?.fillOpacity,
      ),
    }),
    [
      slide.background?.fill,
      slide.background?.fillOpacity,
      slide.height,
      slide.width,
    ],
  );
  const contentStyle = useMemo<CSSProperties>(
    () => ({
      width: slide.width,
      height: slide.height,
      transform: `scale(${thumbnailScale})`,
    }),
    [slide.height, slide.width, thumbnailScale],
  );

  return (
    <div
      className={[
        'office-file-pptx-thumbnail',
        active ? 'office-file-pptx-thumbnail--active' : '',
      ]
        .filter(Boolean)
        .join(' ')}
    >
      <div
        ref={canvasRef}
        className="office-file-pptx-thumbnail__canvas"
        style={canvasStyle}
      >
        <div
          className="office-file-pptx-thumbnail__content"
          style={contentStyle}
        >
          {/* 缩略图复用完整 Slide 渲染，保证背景、图形、表格和图表与主画布一致。 */}
          <PptxSlide
            slide={slide}
            zoom={100}
            renderKey={`thumb-${slide.id}`}
            interactive={false}
          />
        </div>
      </div>
      <div className="office-file-pptx-thumbnail__label">
        {messages.presentation.slide(slide.index)}
      </div>
    </div>
  );
}

export const PptxThumbnail = memo(PptxThumbnailComponent);
