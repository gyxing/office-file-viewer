// PptxViewer 负责 PPTX 预览整体布局，组合左侧缩略图栏和右侧幻灯片视口。
import React, { memo } from 'react';
import type { PresentationDocument } from '../../services/presentation/types';
import { OfficeEmpty } from '../../shell/Empty';
import './index.less';
import { PptxSlideViewport } from './PptxSlideViewport';
import { PptxThumbnailPane } from './PptxThumbnailPane';

/** 定义 PptxViewer 组件可接收的属性。 */
type PptxViewerProps = {
  /** PptxViewerProps 当前关联的标准化文档模型。 */
  document?: PresentationDocument;
  /** 当前选中项在所属集合中的索引。 */
  activeIndex: number;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 在 SelectSlide 事件发生时调用的回调函数。 */
  onSelectSlide: (index: number) => void;
};

/** 渲染 PptxViewerComponent 组件。 */
function PptxViewerComponent({
  document,
  activeIndex,
  zoom,
  onSelectSlide,
}: PptxViewerProps) {
  if (!document?.slides.length) {
    return <OfficeEmpty kind="pptx" />;
  }

  const currentSlide = document.slides[activeIndex];

  return (
    <div className="office-file-pptx-viewer">
      <PptxThumbnailPane
        slides={document.slides}
        activeIndex={activeIndex}
        onSelectSlide={onSelectSlide}
      />
      <PptxSlideViewport
        slide={currentSlide}
        activeIndex={activeIndex}
        zoom={zoom}
      />
    </div>
  );
}

export const PptxViewer = memo(PptxViewerComponent);
