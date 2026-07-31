// PptxViewer 负责 PPTX 预览整体布局，组合左侧缩略图栏和右侧幻灯片视口。
import React, { memo, useMemo } from 'react';
import { getPresentationSource } from '../../services/presentation/presentationSourceRegistry';
import type {
  PresentationDocument,
  PresentationSource,
} from '../../services/presentation/types';
import { OfficeEmpty } from '../../shell/Empty';
import './index.less';
import { PptxSlideViewport } from './PptxSlideViewport';
import { PptxSpeakerNotes } from './PptxSpeakerNotes';
import { PptxThumbnailPane } from './PptxThumbnailPane';
import { usePresentationSource } from './usePresentationSource';

/** PPTX预览器组件属性。 */
type PptxViewerProps = {
  /** 当前处理的标准化文档模型。 */
  document?: PresentationDocument;
  /** 大型演示文稿使用的按页读取数据源。 */
  source?: PresentationSource;
  /** 当前选中项在所属集合中的索引。 */
  activeIndex: number;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 演讲者备注面板当前是否展开。 */
  showSpeakerNotes: boolean;
  /** 在 SelectSlide 事件发生时调用的回调函数。 */
  onSelectSlide: (index: number) => void;
};

/** 渲染PPTX预览器。 */
function PptxViewerComponent({
  document,
  source,
  activeIndex,
  zoom,
  showSpeakerNotes,
  onSelectSlide,
}: PptxViewerProps) {
  const resolvedSource = useMemo(
    () => source ?? (document ? getPresentationSource(document) : undefined),
    [document, source],
  );
  const {
    snapshot,
    slide: currentSlide,
    notes,
    loading,
    error,
    retry,
  } = usePresentationSource(resolvedSource, activeIndex, showSpeakerNotes);

  if (!resolvedSource || !snapshot.slideCount) {
    return <OfficeEmpty kind="pptx" />;
  }

  return (
    <div className="office-file-pptx-viewer">
      <PptxThumbnailPane
        source={resolvedSource}
        snapshot={snapshot}
        activeIndex={activeIndex}
        onSelectSlide={onSelectSlide}
      />
      <div className="office-file-pptx-viewer__workspace">
        <PptxSlideViewport
          slide={currentSlide}
          activeIndex={activeIndex}
          zoom={zoom}
          width={snapshot.width}
          height={snapshot.height}
          loading={loading}
          error={error}
          onRetry={retry}
        />
        {showSpeakerNotes ? (
          <PptxSpeakerNotes
            slideIndex={currentSlide?.index ?? activeIndex + 1}
            notes={notes}
          />
        ) : null}
      </div>
    </div>
  );
}

export const PptxViewer = memo(PptxViewerComponent);
