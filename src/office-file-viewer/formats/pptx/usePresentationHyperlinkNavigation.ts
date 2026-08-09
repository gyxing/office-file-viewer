import { useCallback, useEffect } from 'react';
import type { PresentationSourceSnapshot } from '../../services/presentation/PresentationSource';
import type { OfficeInternalHyperlinkTarget } from '../../shared/hyperlink';
import { useOfficeHyperlinkContext } from '../../shared/hyperlink';

/** 注册演示文稿内部跳转，统一复用现有受控选页状态。 */
export function usePresentationHyperlinkNavigation({
  snapshot,
  activeIndex,
  onSelectSlide,
}: {
  /** 当前文稿的轻量幻灯片快照。 */
  snapshot: PresentationSourceSnapshot;
  /** 当前选中的零基幻灯片索引。 */
  activeIndex: number;
  /** 请求切换当前幻灯片。 */
  onSelectSlide: (index: number) => void;
}) {
  const hyperlinkContext = useOfficeHyperlinkContext();
  const navigate = useCallback(
    (target: OfficeInternalHyperlinkTarget) => {
      if (target.family !== 'presentation' || !snapshot.slideCount) {
        return false;
      }
      let nextIndex = target.slideId
        ? snapshot.slides.findIndex((slide) => slide.id === target.slideId)
        : target.slideIndex;
      if (target.action === 'first') nextIndex = 0;
      if (target.action === 'last') nextIndex = snapshot.slideCount - 1;
      if (target.action === 'next') {
        nextIndex = Math.min(snapshot.slideCount - 1, activeIndex + 1);
      }
      if (target.action === 'previous') {
        nextIndex = Math.max(0, activeIndex - 1);
      }
      if (
        nextIndex === undefined ||
        nextIndex < 0 ||
        nextIndex >= snapshot.slideCount
      ) {
        return false;
      }
      onSelectSlide(nextIndex);
      return true;
    },
    [activeIndex, onSelectSlide, snapshot],
  );

  useEffect(() => {
    if (!hyperlinkContext) return undefined;
    return hyperlinkContext.registerNavigator('presentation', navigate);
  }, [hyperlinkContext, navigate]);
}
