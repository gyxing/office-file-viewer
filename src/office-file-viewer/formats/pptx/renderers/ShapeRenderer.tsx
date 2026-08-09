// ShapeRenderer 渲染 PPTX 基础图形和自定义路径图形。
import React, { memo } from 'react';
import type { ShapeElement } from '../../../services/pptx/types';
import { useOfficeHyperlink } from '../../../shared/hyperlink';
import {
  colorWithOpacity,
  gradientToSvgEndpoints,
  isGradientPaint,
  paintToCss,
} from './paint';
import { buildRendererId } from './renderIds';

/** 形状渲染器组件属性。 */
type ShapeRendererProps = {
  /** 当前处理或渲染的演示文稿元素。 */
  element: ShapeElement;
  /** 内容变化时用于刷新渲染结果的键。 */
  renderKey: string;
  /** 是否允许当前形状响应链接交互。 */
  interactive: boolean;
};

function lineStyle(dash?: string) {
  if (!dash || dash === 'solid') return 'solid';
  if (dash.includes('dot')) return 'dotted';
  return 'dashed';
}

/** 渲染形状渲染器。 */
function ShapeRendererComponent({
  element,
  renderKey,
  interactive,
}: ShapeRendererProps) {
  const hyperlinkProps = useOfficeHyperlink<HTMLDivElement>({
    hyperlink: element.hyperlink,
    source: {
      type: element.hyperlinkSourceType ?? 'shape',
      id: element.id,
    },
    interactive,
  });
  const fillPaint = element.fill;
  const isGradientFill = isGradientPaint(fillPaint);
  const gradientId = isGradientFill
    ? buildRendererId(renderKey, element.id, 'fill-gradient')
    : undefined;
  const radius =
    element.shape === 'ellipse'
      ? '50%'
      : element.shape === 'roundRect'
      ? Math.min(element.width, element.height) * (element.borderRadius ?? 0.12)
      : 0;
  const isLineShape = element.shape === 'line';
  const isVectorShape = Boolean(element.path);
  const shadow = element.shadow
    ? `${element.shadow.offsetX ?? 0}px ${
        element.shadow.offsetY ?? 0
      }px ${Math.max(0, element.shadow.blur ?? 0)}px ${colorWithOpacity(
        element.shadow.color ?? 'rgba(0,0,0,0.18)',
        element.shadow.opacity,
      )}`
    : undefined;

  return (
    <div
      {...hyperlinkProps}
      style={{
        position: 'absolute',
        left: element.x,
        top: element.y,
        width: isLineShape ? Math.max(1, element.width) : element.width,
        height: isLineShape ? Math.max(1, element.height) : element.height,
        background:
          isVectorShape || isLineShape
            ? 'transparent'
            : paintToCss(element.fill, element.fillOpacity) ?? 'transparent',
        border:
          isVectorShape || isLineShape
            ? undefined
            : element.stroke
            ? `${element.strokeWidth ?? 1}px ${lineStyle(
                element.strokeDash,
              )} ${colorWithOpacity(element.stroke, element.strokeOpacity)}`
            : undefined,
        borderRadius: radius,
        boxShadow: shadow,
        transform: element.rotate
          ? `rotate(${element.rotate}deg)${element.flipH ? ' scaleX(-1)' : ''}${
              element.flipV ? ' scaleY(-1)' : ''
            }`
          : `${element.flipH ? 'scaleX(-1)' : ''}${
              element.flipV ? ' scaleY(-1)' : ''
            }`.trim(),
        transformOrigin: 'center center',
        overflow: 'visible',
      }}
    >
      {isVectorShape || isLineShape ? (
        <svg
          viewBox={
            element.viewBox ??
            `0 0 ${Math.max(1, element.width)} ${Math.max(1, element.height)}`
          }
          preserveAspectRatio="none"
          style={{
            position: 'absolute',
            inset: 0,
            width: '100%',
            height: '100%',
            overflow: 'visible',
          }}
        >
          {isGradientFill ? (
            <defs>
              <linearGradient
                id={gradientId}
                {...gradientToSvgEndpoints(fillPaint.angle)}
                gradientUnits="objectBoundingBox"
              >
                {fillPaint.stops.map((stop, index) => (
                  <stop
                    key={index}
                    offset={`${stop.offset * 100}%`}
                    stopColor={stop.color}
                  />
                ))}
              </linearGradient>
            </defs>
          ) : null}
          {isLineShape ? (
            <line
              x1="0"
              y1="0"
              x2={Math.max(0, element.width)}
              y2={Math.max(0, element.height)}
              stroke={
                element.stroke
                  ? colorWithOpacity(element.stroke, element.strokeOpacity) ??
                    element.stroke
                  : 'none'
              }
              strokeOpacity={element.strokeOpacity}
              strokeWidth={element.strokeWidth ?? 1}
              strokeDasharray={
                element.strokeDash && element.strokeDash !== 'solid'
                  ? element.strokeDash
                  : undefined
              }
              strokeLinecap="round"
              vectorEffect="non-scaling-stroke"
            />
          ) : (
            <path
              d={element.path ?? ''}
              fill={
                isGradientFill
                  ? `url(#${gradientId})`
                  : element.fill
                  ? colorWithOpacity(
                      element.fill as string,
                      element.fillOpacity,
                    ) ?? 'none'
                  : 'none'
              }
              fillOpacity={isGradientFill ? undefined : element.fillOpacity}
              stroke={
                element.stroke
                  ? colorWithOpacity(element.stroke, element.strokeOpacity) ??
                    element.stroke
                  : 'none'
              }
              strokeOpacity={element.strokeOpacity}
              strokeWidth={element.strokeWidth ?? 1}
              strokeDasharray={
                element.strokeDash && element.strokeDash !== 'solid'
                  ? element.strokeDash
                  : undefined
              }
              vectorEffect="non-scaling-stroke"
            />
          )}
        </svg>
      ) : null}
    </div>
  );
}

export const ShapeRenderer = memo(ShapeRendererComponent);
