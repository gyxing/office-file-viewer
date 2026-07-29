import { Button } from 'antd';
import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';

type PresentationSlideStateProps = {
  width: number;
  height: number;
  error?: Error;
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
        <Button size="small" onClick={onRetry}>
          {messages.lazyContent.retry}
        </Button>
      ) : null}
    </div>
  );
}

export const PresentationSlideState = memo(PresentationSlideStateComponent);
