// UnsupportedRenderer 渲染尚未支持的 PPTX 元素占位，避免元素静默丢失。
import React, { memo } from 'react';
import type { UnsupportedElement } from '../../../services/pptx/types';

/** 不支持渲染器组件属性。 */
type UnsupportedRendererProps = {
  /** 当前处理或渲染的演示文稿元素。 */
  element: UnsupportedElement;
};

/** 渲染暂不支持的幻灯片元素占位。 */
function UnsupportedRendererComponent({ element }: UnsupportedRendererProps) {
  return (
    <div
      style={{
        position: 'absolute',
        left: element.x,
        top: element.y,
        width: element.width,
        height: element.height,
        border: '1px dashed #d92d20',
        color: '#d92d20',
        fontSize: 12,
      }}
    >
      {element.reason}
    </div>
  );
}

export const UnsupportedRenderer = memo(UnsupportedRendererComponent);
