// PptxThumbnailPane 渲染幻灯片缩略图列表，并负责切换当前页。
import React, { memo, useCallback } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type { SlideModel } from '../../services/pptx/types';
import { PptxThumbnail } from './PptxThumbnail';

/** 定义 PptxThumbnailPane 组件可接收的属性。 */
type PptxThumbnailPaneProps = {
  /** PptxThumbnailPaneProps 包含的 slides 有序集合。 */
  slides: SlideModel[];
  /** 当前选中项在所属集合中的索引。 */
  activeIndex: number;
  /** 在 SelectSlide 事件发生时调用的回调函数。 */
  onSelectSlide: (index: number) => void;
};

/** 渲染 PptxThumbnailPaneComponent 组件。 */
function PptxThumbnailPaneComponent({
  slides,
  activeIndex,
  onSelectSlide,
}: PptxThumbnailPaneProps) {
  const messages = useOfficeFileViewerMessages();
  const handleSelect = useCallback(
    (index: number) => {
      onSelectSlide(index);
    },
    [onSelectSlide],
  );

  return (
    <aside className="office-file-pptx-viewer__sidebar">
      <div className="office-file-pptx-viewer__sidebar-header">
        <div className="office-file-pptx-viewer__slide-count">
          {messages.presentation.slideCount(slides.length)}
        </div>
      </div>
      <div className="office-file-pptx-viewer__thumbnail-list">
        {slides.map((slide, index) => (
          <button
            key={slide.id}
            type="button"
            className="office-file-pptx-viewer__thumbnail-button"
            onClick={() => handleSelect(index)}
          >
            <PptxThumbnail slide={slide} active={index === activeIndex} />
          </button>
        ))}
      </div>
    </aside>
  );
}

export const PptxThumbnailPane = memo(PptxThumbnailPaneComponent);
