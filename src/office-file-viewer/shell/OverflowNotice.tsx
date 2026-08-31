import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../locale';
import type { OfficeOverflowAxis } from './useOfficeOverflowNotice';

/** 在内容确实超出视口时显示可滚动方向提示。 */
function OverflowNoticeComponent({
  axis,
}: {
  /** 当前内容超出的方向。 */
  axis?: OfficeOverflowAxis;
}) {
  const messages = useOfficeFileViewerMessages();
  if (!axis) return null;
  const message =
    axis === 'both'
      ? messages.viewport.bothOverflow
      : axis === 'horizontal'
      ? messages.viewport.horizontalOverflow
      : messages.viewport.verticalOverflow;
  return (
    <div
      className="office-file-overflow-notice"
      data-office-overflow-notice="true"
      data-office-overflow-axis={axis}
      role="status"
      aria-live="polite"
    >
      {message}
    </div>
  );
}

export const OfficeOverflowNotice = memo(OverflowNoticeComponent);
