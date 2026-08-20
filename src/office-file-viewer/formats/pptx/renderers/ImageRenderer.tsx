// ImageRenderer 渲染 PPTX 图片元素，并处理裁剪、旋转和翻转。
import React, { memo, useMemo } from 'react';
import type { ImageElement } from '../../../services/pptx/types';
import {
  useOfficeResourceUrl,
  type OfficeResourceSource,
} from '../../../services/resource-store';
import { useOfficeHyperlink } from '../../../shared/hyperlink';

/** 图片渲染器组件属性。 */
type ImageRendererProps = {
  /** 当前处理或渲染的演示文稿元素。 */
  element: ImageElement;
  /** 是否允许当前图片响应链接交互。 */
  interactive: boolean;
};

/** 渲染图片渲染器。 */
function ImageRendererComponent({ element, interactive }: ImageRendererProps) {
  const hyperlinkProps = useOfficeHyperlink<HTMLDivElement>({
    hyperlink: element.hyperlink,
    source: { type: 'image', id: element.id },
    interactive,
  });
  const source = useMemo<OfficeResourceSource>(
    () =>
      typeof element.src === 'string'
        ? { kind: 'url', url: element.src }
        : element.src,
    [element.src],
  );
  const resource = useOfficeResourceUrl(source);
  const left = element.crop?.left ?? 0;
  const top = element.crop?.top ?? 0;
  const right = element.crop?.right ?? 0;
  const bottom = element.crop?.bottom ?? 0;
  const visibleWidth = Math.max(0.01, 1 - left - right);
  const visibleHeight = Math.max(0.01, 1 - top - bottom);

  return (
    <div
      {...hyperlinkProps}
      data-office-presentation-element-id={element.id}
      style={{
        position: 'absolute',
        left: element.x,
        top: element.y,
        width: element.width,
        height: element.height,
        overflow: 'hidden',
        transform: [
          element.rotate ? `rotate(${element.rotate}deg)` : '',
          element.flipH ? 'scaleX(-1)' : '',
          element.flipV ? 'scaleY(-1)' : '',
        ]
          .filter(Boolean)
          .join(' '),
        transformOrigin: 'center center',
        pointerEvents: interactive && element.hyperlink ? 'auto' : 'none',
      }}
    >
      <img
        alt={element.alt ?? ''}
        src={resource.url}
        loading="lazy"
        decoding="async"
        style={{
          position: 'absolute',
          left: `${-(left / visibleWidth) * 100}%`,
          top: `${-(top / visibleHeight) * 100}%`,
          width: `${100 / visibleWidth}%`,
          height: `${100 / visibleHeight}%`,
          objectFit: 'fill',
          display: 'block',
        }}
      />
    </div>
  );
}

export const ImageRenderer = memo(ImageRendererComponent);
