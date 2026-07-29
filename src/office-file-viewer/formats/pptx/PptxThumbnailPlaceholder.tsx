import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type { PresentationSlideDescriptor } from '../../services/presentation/types';

type PptxThumbnailPlaceholderProps = {
  descriptor: PresentationSlideDescriptor;
  aspectRatio: number;
  active: boolean;
};

/** 为尚未读取或读取失败的幻灯片保留稳定缩略图尺寸。 */
function PptxThumbnailPlaceholderComponent({
  descriptor,
  aspectRatio,
  active,
}: PptxThumbnailPlaceholderProps) {
  const messages = useOfficeFileViewerMessages();
  const failed = descriptor.status === 'error';

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
        className="office-file-pptx-thumbnail__placeholder"
        style={{ aspectRatio }}
      >
        <span>
          {failed
            ? messages.lazyContent.slideLoadFailed
            : messages.lazyContent.loading}
        </span>
        {failed ? (
          <span className="office-file-pptx-thumbnail__retry">
            {messages.lazyContent.retry}
          </span>
        ) : null}
      </div>
      <div className="office-file-pptx-thumbnail__label">
        {messages.presentation.slide(descriptor.index)}
      </div>
    </div>
  );
}

export const PptxThumbnailPlaceholder = memo(PptxThumbnailPlaceholderComponent);
