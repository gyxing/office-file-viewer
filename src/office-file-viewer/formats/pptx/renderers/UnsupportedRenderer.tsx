// UnsupportedRenderer 渲染尚未支持的 PPTX 元素占位，避免元素静默丢失。
import React, { memo } from 'react';
import type { UnsupportedElement } from '../../../services/pptx/types';

/** 定义 UnsupportedRenderer 组件可接收的属性。 */
type UnsupportedRendererProps = {
  /** UnsupportedRendererProps 当前负责渲染的演示文稿元素模型。 */
  element: UnsupportedElement;
};

/** 渲染 UnsupportedRendererComponent 组件。 */
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
