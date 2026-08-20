// PptxViewer 负责 PPTX 预览整体布局，组合左侧缩略图栏和右侧幻灯片视口。
import React, { memo, useEffect, useMemo, useRef } from 'react';
import type { OfficeFileViewerPreviewState } from '../../services/parsing/internalTypes';
import type { OfficeFileViewerPresentationMediaOptions } from '../../services/presentation/mediaTypes';
import { PresentationAnnotationSource } from '../../services/presentation/PresentationAnnotationSource';
import { getPresentationSource } from '../../services/presentation/presentationSourceRegistry';
import type {
  OfficeFileViewerPresentationTransitions,
  PresentationNavigationIntent,
} from '../../services/presentation/transitionTypes';
import { useOfficeAnnotationSourceRegistration } from '../../shared/annotations';
import { OfficePreviewEmpty } from '../common/OfficePreviewEmpty';
import { useOfficeSearchProviderRegistration } from '../search/OfficeSearchContext';
import './index.less';
import { PptxSlideViewport } from './PptxSlideViewport';
import { PptxSpeakerNotes } from './PptxSpeakerNotes';
import { PptxThumbnailPane } from './PptxThumbnailPane';
import { usePresentationAnnotationNavigation } from './usePresentationAnnotationNavigation';
import { usePresentationHyperlinkNavigation } from './usePresentationHyperlinkNavigation';
import { usePresentationSearchNavigation } from './usePresentationSearchNavigation';
import { usePresentationSource } from './usePresentationSource';

/** 演示文稿 Viewer 可以消费的物化或按需预览。 */
type PresentationPreview = Extract<
  OfficeFileViewerPreviewState,
  { previewKind: 'ppt' | 'pptx' }
>;

/** PPTX预览器组件属性。 */
type PptxViewerProps = {
  /** 当前演示文稿的物化或按需预览。 */
  preview: PresentationPreview;
  /** 当前选中项在所属集合中的索引。 */
  activeIndex: number;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 演讲者备注面板当前是否展开。 */
  showSpeakerNotes: boolean;
  /** 在 SelectSlide 事件发生时调用的回调函数。 */
  onSelectSlide: (index: number) => void;
  /** 演示文稿媒体读取配置。 */
  mediaOptions?: false | OfficeFileViewerPresentationMediaOptions;
  /** 是否按源文件播放页级切换。 */
  transitions: OfficeFileViewerPresentationTransitions;
  /** 最近一次工具栏翻页产生的切换意图。 */
  transitionIntent?: PresentationNavigationIntent;
};

/** 渲染PPTX预览器。 */
function PptxViewerComponent({
  preview,
  activeIndex,
  zoom,
  showSpeakerNotes,
  onSelectSlide,
  mediaOptions,
  transitions,
  transitionIntent,
}: PptxViewerProps) {
  const resolvedSource = useMemo(
    () =>
      preview.mode === 'source'
        ? preview.source
        : getPresentationSource(preview.model.document),
    [preview],
  );
  const viewerRef = useRef<HTMLDivElement>(null);
  useOfficeSearchProviderRegistration(resolvedSource?.searchProvider);
  const {
    snapshot,
    slide: currentSlide,
    notes,
    loading,
    error,
    retry,
  } = usePresentationSource(resolvedSource, activeIndex, showSpeakerNotes);
  const annotationSource = useMemo(
    () =>
      resolvedSource
        ? new PresentationAnnotationSource(resolvedSource)
        : undefined,
    [resolvedSource],
  );
  useEffect(() => () => annotationSource?.dispose(), [annotationSource]);
  useOfficeAnnotationSourceRegistration(annotationSource);
  usePresentationAnnotationNavigation({ viewerRef, onSelectSlide });
  usePresentationHyperlinkNavigation({ snapshot, activeIndex, onSelectSlide });
  usePresentationSearchNavigation({
    snapshot,
    onSelectSlide,
    viewerRef,
  });

  if (!resolvedSource || !snapshot.slideCount) {
    return <OfficePreviewEmpty kind={preview.previewKind} />;
  }

  return (
    <div ref={viewerRef} className="office-file-pptx-viewer">
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
          mediaOptions={mediaOptions}
          transitions={transitions}
          transitionIntent={transitionIntent}
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
