// DocxShape 渲染 DOCX 行内形状，支持矩形、椭圆、线条、自定义路径和形状内文字。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { DocxInline } from '../../services/docx/types';
import { useOfficeHyperlink } from '../../shared/hyperlink';
import { DocxShapeItemRenderer } from './DocxShapeItemRenderer';
import { calculatePositionStyle } from './positionUtils';

/** DOCX形状组件属性。 */
type DocxShapeProps = {
  /** 当前负责渲染的行内内容模型。 */
  inline: Extract<
    DocxInline,
    {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'shape';
    }
  >;
};

// 自定义变量把解析后的形状尺寸交给 Less，保持定位样式的类型约束。
/** DOCX 形状渲染样式。 */
type DocxShapeStyle = CSSProperties & {
  /** 传递给样式表的形状宽度 CSS 自定义属性。 */
  '--office-file-docx-shape-width': string;
  /** 传递给样式表的形状高度 CSS 自定义属性。 */
  '--office-file-docx-shape-height': string;
};

/** 渲染DOCX形状。 */
function DocxShapeComponent({ inline }: DocxShapeProps) {
  const shape = inline.shape;
  const positionStyle = calculatePositionStyle(shape.position);
  const hyperlinkProps = useOfficeHyperlink<HTMLSpanElement>({
    hyperlink: shape.hyperlink,
    source: { type: 'shape', id: shape.id },
  });

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
    <span
      {...hyperlinkProps}
      className="office-file-docx-shape"
      style={shapeStyle}
    >
      {shape.items.map((item) => (
        <DocxShapeItemRenderer key={item.id} item={item} />
      ))}
    </span>
  );
}

export const DocxShape = memo(DocxShapeComponent);
