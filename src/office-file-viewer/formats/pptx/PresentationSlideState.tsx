import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import { OfficeButton } from '../../shared/ui/OfficeButton';

/** 幻灯片加载状态组件属性。 */
type PresentationSlideStateProps = {
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 当前操作产生的错误；未提供表示没有错误。 */
  error?: Error;
  /** 用户请求重新加载时触发的回调。 */
  onRetry: () => void;
};

/** 在主视口内显示保持幻灯片宽高比的加载或单页失败状态。 */
function PresentationSlideStateComponent({
  width,
  height,
  error,
  onRetry,
}: PresentationSlideStateProps) {
  const messages = useOfficeFileViewerMessages();

  return (
    <div
      className="office-file-pptx-slide-state"
      style={{ width, aspectRatio: `${width} / ${height}` }}
    >
      <span>
        {error
          ? messages.lazyContent.slideLoadFailed
          : messages.lazyContent.loading}
      </span>
      {error ? (
        <OfficeButton size="small" onClick={onRetry}>
          {messages.lazyContent.retry}
        </OfficeButton>
      ) : null}
    </div>
  );
}

export const PresentationSlideState = memo(PresentationSlideStateComponent);
