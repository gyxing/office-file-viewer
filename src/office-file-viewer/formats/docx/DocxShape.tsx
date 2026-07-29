// DocxShape 渲染 DOCX 行内形状，支持矩形、椭圆、线条、自定义路径和形状内文字。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { DocxInline } from '../../services/docx/types';
import { DocxShapeItemRenderer } from './DocxShapeItemRenderer';
import { calculatePositionStyle } from './positionUtils';

/** 定义 DocxShape 组件可接收的属性。 */
type DocxShapeProps = {
  /** DocxShapeProps 当前负责渲染的行内内容模型。 */
  inline: Extract<
    DocxInline,
    {
      /** 用于区分 DocxShapeProps 不同结构分支的类型标识。 */ type: 'shape';
    }
  >;
};

// 自定义变量把解析后的形状尺寸交给 Less，保持定位样式的类型约束。
/** 描述 DOCX 渲染使用的样式参数。 */
type DocxShapeStyle = CSSProperties & {
  /** DocxShapeStyle 的 --office-file-docx-shape-width 文本值。 */
  '--office-file-docx-shape-width': string;
  /** DocxShapeStyle 的 --office-file-docx-shape-height 文本值。 */
  '--office-file-docx-shape-height': string;
};

/** 渲染 DocxShapeComponent 组件。 */
function DocxShapeComponent({ inline }: DocxShapeProps) {
  const shape = inline.shape;
  const positionStyle = calculatePositionStyle(shape.position);

  const shapeStyle = useMemo<DocxShapeStyle>(() => {
    return {
      '--office-file-docx-shape-width': `${shape.width}px`,
      '--office-file-docx-shape-height': `${shape.height}px`,
      ...positionStyle,
      zIndex: positionStyle.zIndex,
      maxWidth: shape.position ? 'none' : undefined,
      margin: shape.position ? 0 : undefined,
    };
  }, [positionStyle, shape.height, shape.position, shape.width]);
  return (
    <span className="office-file-docx-shape" style={shapeStyle}>
      {shape.items.map((item) => (
        <DocxShapeItemRenderer key={item.id} item={item} />
      ))}
    </span>
  );
}

export const DocxShape = memo(DocxShapeComponent);
