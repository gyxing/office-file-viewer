// PptxSlide 按 PPTX 解析模型渲染单页幻灯片背景、图形、文本、图片、表格和图表。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { SlideElement, SlideModel } from '../../services/pptx/types';
import {
  useOfficeResourceUrl,
  type OfficeResourceSource,
} from '../../services/resource-store';
import { OfficeChartView } from '../../shared/chart/OfficeChartView';
import { useOfficeHyperlink } from '../../shared/hyperlink';
import { ImageRenderer } from './renderers/ImageRenderer';
import { colorWithOpacity } from './renderers/paint';
import { ShapeRenderer } from './renderers/ShapeRenderer';
import { TableRenderer } from './renderers/TableRenderer';
import { TextRenderer } from './renderers/TextRenderer';
import { UnsupportedRenderer } from './renderers/UnsupportedRenderer';

/** PPTX幻灯片组件属性。 */
type PptxSlideProps = {
  /** 当前处理或展示的幻灯片。 */
  slide: SlideModel;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 内容变化时用于刷新渲染结果的键。 */
  renderKey?: string;
  /** 是否允许幻灯片内部对象响应链接交互。 */
  interactive?: boolean;
  /** 主视口中的零基幻灯片索引；缩略图不传入以避免重复高亮。 */
  searchSlideIndex?: number;
};

const ChartFrame = memo(function ChartFrame({
  element,
  interactive,
}: {
  /** 当前负责渲染的演示文稿元素。 */
  element: Extract<
    SlideElement,
    {
      /** 固定为 `chart`，用于区分联合类型分支。 */
      type: 'chart';
    }
  >;
  /** 是否允许当前图表响应链接交互。 */
  interactive: boolean;
}) {
  const hyperlinkProps = useOfficeHyperlink<HTMLDivElement>({
    hyperlink: element.hyperlink,
    source: { type: 'shape', id: element.id },
    interactive,
  });
  const snapshot = useOfficeResourceUrl(element.snapshotSource);
  const chart = useMemo(
    () =>
      snapshot.url
        ? { ...element.chart, snapshotSrc: snapshot.url }
        : element.chart,
    [element.chart, snapshot.url],
  );
  const frameStyle = useMemo<CSSProperties>(
    () => ({
      left: element.x,
      top: element.y,
      width: element.width,
      height: element.height,
    }),
    [element.height, element.width, element.x, element.y],
  );

  return (
    <div
      {...hyperlinkProps}
      className="office-file-pptx-chart-frame"
      style={frameStyle}
    >
      <OfficeChartView
        chart={chart}
        width={element.width}
        height={element.height}
      />
    </div>
  );
});

/** 渲染PPTX幻灯片。 */
function PptxSlideComponent({
  slide,
  zoom,
  renderKey,
  interactive = true,
  searchSlideIndex,
}: PptxSlideProps) {
  const scale = zoom / 100;
  const backgroundSource = useMemo<OfficeResourceSource | undefined>(
    () =>
      typeof slide.background?.imageRef === 'string'
        ? { kind: 'url', url: slide.background.imageRef }
        : slide.background?.imageRef,
    [slide.background?.imageRef],
  );
  const backgroundResource = useOfficeResourceUrl(backgroundSource);
  // renderKey 会参与 SVG 渐变 id，缩略图和主画布同时渲染同一页时必须保持 id 隔离。
  const slideRenderKey = renderKey ?? `slide-${slide.id}`;
  const slideStyle = useMemo<CSSProperties>(
    () => ({
      width: slide.width,
      height: slide.height,
      minWidth: slide.width,
      minHeight: slide.height,
      transform: `scale(${scale})`,
    }),
    [scale, slide.height, slide.width],
  );
  const backgroundStyle = useMemo<CSSProperties>(
    () => ({
      // 避免 background 简写重置样式表中的 cover，图片背景必须按整页尺寸铺满。
      backgroundColor: colorWithOpacity(
        slide.background?.fill ?? '#fff',
        slide.background?.fillOpacity,
      ),
      backgroundImage: backgroundResource.url
        ? `url(${backgroundResource.url})`
        : undefined,
    }),
    [
      slide.background?.fill,
      slide.background?.fillOpacity,
      backgroundResource.url,
    ],
  );

  return (
    <div className="office-file-pptx-slide" style={slideStyle}>
      <div
        className="office-file-pptx-slide__background"
        style={backgroundStyle}
      />
      <div className="office-file-pptx-slide__elements">
        {slide.elements.map((element) => {
          switch (element.type) {
            case 'text':
              return (
                <TextRenderer
                  key={element.id}
                  element={element}
                  renderKey={slideRenderKey}
                  interactive={interactive}
                  searchSlideIndex={searchSlideIndex}
                />
              );
            case 'shape':
              return (
                <ShapeRenderer
                  key={element.id}
                  element={element}
                  renderKey={slideRenderKey}
                  interactive={interactive}
                />
              );
            case 'image':
              return (
                <ImageRenderer
                  key={element.id}
                  element={element}
                  interactive={interactive}
                />
              );
            case 'table':
              return (
                <TableRenderer
                  key={element.id}
                  element={element}
                  interactive={interactive}
                  searchSlideIndex={searchSlideIndex}
                />
              );
            case 'chart':
              return (
                <ChartFrame
                  key={element.id}
                  element={element}
                  interactive={interactive}
                />
              );
            case 'unsupported':
              return <UnsupportedRenderer key={element.id} element={element} />;
            default:
              return null;
          }
        })}
      </div>
    </div>
  );
}

export const PptxSlide = memo(PptxSlideComponent);
