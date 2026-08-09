// TextRenderer 渲染 PPTX 文本框，并处理文本框形状、渐变填充、段落和 run 样式。
import type { CSSProperties } from 'react';
import React, { memo } from 'react';
import type {
  TextElement,
  TextRun,
  TextStyle,
} from '../../../services/pptx/types';
import { useOfficeHyperlink } from '../../../shared/hyperlink';
import {
  colorWithOpacity,
  gradientToSvgEndpoints,
  isGradientPaint,
  paintToCss,
} from './paint';
import { buildRendererId } from './renderIds';

/** 文本渲染器组件属性。 */
type TextRendererProps = {
  /** 当前处理或渲染的演示文稿元素。 */
  element: TextElement;
  /** 内容变化时用于刷新渲染结果的键。 */
  renderKey: string;
  /** 是否允许当前文本框及文字片段响应链接交互。 */
  interactive: boolean;
};

function shadowToCss(element: TextElement) {
  if (!element.shadow) return undefined;
  return `${element.shadow.offsetX ?? 0}px ${
    element.shadow.offsetY ?? 0
  }px ${Math.max(0, element.shadow.blur ?? 0)}px ${colorWithOpacity(
    element.shadow.color ?? '#000000',
    element.shadow.opacity ?? 0.18,
  )}`;
}

function radiusToPx(element: TextElement) {
  if (element.shape === 'ellipse') return '50%';
  if (element.shape !== 'roundRect') return 0;
  const ratio = element.borderRadius ?? 0.12;
  return Math.min(element.width, element.height) * ratio;
}

function textDecoration(style: NonNullable<TextElement['boxStyle']>) {
  const parts: string[] = [];
  if (style.underline) parts.push('underline');
  if (style.strike && style.strike !== 'none') parts.push('line-through');
  return parts.length ? parts.join(' ') : 'none';
}

/** 把文本纯色或渐变填充转换为浏览器字形样式。 */
function resolveTextPaintStyle(
  fill: TextStyle['textFill'],
  fallbackColor: TextStyle['color'],
  opacity?: number,
): CSSProperties {
  if (isGradientPaint(fill)) {
    return {
      color: 'transparent',
      backgroundImage: paintToCss(fill),
      backgroundClip: 'text',
      WebkitBackgroundClip: 'text',
      WebkitTextFillColor: 'transparent',
    };
  }
  return { color: colorWithOpacity(fill ?? fallbackColor, opacity) };
}

function lineStyle(dash?: string) {
  if (!dash || dash === 'solid') return 'solid';
  if (dash.includes('dot')) return 'dotted';
  return 'dashed';
}

/** 渲染单个文字片段，使链接 Hook 不违反 React 循环调用约束。 */
function TextRunRenderer({
  run,
  boxStyle,
  sourceId,
  interactive,
}: {
  /** 当前文字片段。 */
  run: TextRun;
  /** 文本框继承的基础样式。 */
  boxStyle: TextStyle;
  /** 当前片段在文稿中的稳定来源标识。 */
  sourceId: string;
  /** 是否允许当前文字片段响应链接交互。 */
  interactive: boolean;
}) {
  const hyperlinkProps = useOfficeHyperlink<HTMLSpanElement>({
    hyperlink: run.hyperlink,
    source: { type: 'text', id: sourceId },
    interactive,
  });
  const runStyle = run.style ?? {};
  const underline = runStyle.underline ?? boxStyle.underline;
  const strike = runStyle.strike ?? boxStyle.strike;
  return (
    <span
      {...hyperlinkProps}
      style={{
        ...resolveTextPaintStyle(
          runStyle.textFill ?? boxStyle.textFill,
          runStyle.color ?? boxStyle.color ?? '#172033',
          runStyle.opacity ?? boxStyle.opacity,
        ),
        fontFamily: runStyle.fontFamily ?? boxStyle.fontFamily,
        fontSize: runStyle.fontSize ?? boxStyle.fontSize,
        fontWeight: runStyle.bold ?? boxStyle.bold ? 700 : 400,
        fontStyle: runStyle.italic ?? boxStyle.italic ? 'italic' : 'normal',
        textDecoration:
          [
            underline ? 'underline' : '',
            strike && strike !== 'none' ? 'line-through' : '',
          ]
            .filter(Boolean)
            .join(' ') || 'none',
        textTransform:
          runStyle.allCaps ?? boxStyle.allCaps ? 'uppercase' : undefined,
        fontVariant:
          runStyle.smallCaps ?? boxStyle.smallCaps ? 'small-caps' : undefined,
        verticalAlign:
          runStyle.baseline && runStyle.baseline > 0
            ? 'super'
            : runStyle.baseline && runStyle.baseline < 0
            ? 'sub'
            : undefined,
        letterSpacing: runStyle.charSpace ?? boxStyle.charSpace ?? 0,
      }}
    >
      {run.text}
    </span>
  );
}

/** 渲染文本渲染器。 */
function TextRendererComponent({
  element,
  renderKey,
  interactive,
}: TextRendererProps) {
  const style = element.boxStyle ?? {};
  const hyperlinkProps = useOfficeHyperlink<HTMLDivElement>({
    hyperlink: element.hyperlink,
    source: {
      type: element.hyperlinkSourceType ?? 'shape',
      id: element.id,
    },
    interactive,
  });
  const isVectorShape = Boolean(element.path);
  const fillPaint = element.fill;
  const isGradientFill = isGradientPaint(fillPaint);
  // 同一页可能在缩略图和主视口同时出现，SVG defs id 需要带 renderKey 防止互相引用错。
  const gradientId = isGradientFill
    ? buildRendererId(renderKey, element.id, 'fill-gradient')
    : undefined;

  return (
    <div
      {...hyperlinkProps}
      style={{
        position: 'absolute',
        left: element.x,
        top: element.y,
        width: element.width,
        height: element.height,
        ...resolveTextPaintStyle(
          style.textFill,
          style.color ?? '#172033',
          style.opacity,
        ),
        fontFamily: style.fontFamily,
        fontSize: style.fontSize,
        fontWeight: style.bold ? 700 : 400,
        fontStyle: style.italic ? 'italic' : 'normal',
        textDecoration: textDecoration(style),
        textTransform: style.allCaps ? 'uppercase' : undefined,
        fontVariant: style.smallCaps ? 'small-caps' : undefined,
        writingMode: style.writingMode,
        whiteSpace: 'pre-wrap',
        lineHeight: style.lineHeight ?? 1.15,
        background: isVectorShape
          ? undefined
          : paintToCss(element.fill, element.fillOpacity),
        border:
          !isVectorShape && element.stroke
            ? `${element.strokeWidth ?? 1}px ${lineStyle(
                element.strokeDash,
              )} ${colorWithOpacity(element.stroke, element.strokeOpacity)}`
            : undefined,
        borderRadius: radiusToPx(element),
        boxShadow: shadowToCss(element),
        boxSizing: 'border-box',
        overflow: 'hidden',
        display: 'flex',
        flexDirection: 'column',
        justifyContent:
          style.verticalAlign === 'bottom'
            ? 'flex-end'
            : style.verticalAlign === 'middle'
            ? 'center'
            : 'flex-start',
        paddingLeft: style.marginLeft ?? 0,
        paddingRight: style.marginRight ?? 0,
        paddingTop: style.marginTop ?? 0,
        paddingBottom: style.marginBottom ?? 0,
        letterSpacing: style.charSpace ?? 0,
        // PowerPoint 不会在连续数字或英文单词中间强制断行，窄文本框应保留原词并裁切。
        wordBreak: 'normal',
        overflowWrap: 'normal',
      }}
    >
      {isVectorShape ? (
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
            zIndex: 0,
            pointerEvents: 'none',
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
        </svg>
      ) : null}
      {element.paragraphs.map((paragraph, paragraphIndex) => (
        <div
          key={paragraphIndex}
          style={{
            position: 'relative',
            zIndex: 1,
            textAlign: paragraph.style?.align ?? style.align ?? 'left',
            lineHeight: paragraph.style?.lineHeight ?? style.lineHeight ?? 1.2,
            marginTop: paragraph.style?.spaceBefore ?? 0,
            marginBottom: paragraph.style?.spaceAfter ?? 0,
            paddingLeft: `${
              (paragraph.style?.marginLeft ?? 0) +
              (paragraph.bullet && !paragraph.bullet.none ? 18 : 0)
            }px`,
            textIndent: paragraph.style?.textIndent
              ? `${paragraph.style.textIndent}px`
              : undefined,
            display: 'block',
            whiteSpace: 'inherit',
          }}
        >
          {paragraph.bullet && !paragraph.bullet.none ? (
            <span
              style={{
                display: 'inline-block',
                color: colorWithOpacity(
                  paragraph.bullet.color ?? style.color,
                  style.opacity,
                ),
                fontSize: paragraph.bullet.size ?? style.fontSize,
                marginRight: 6,
                width: 12,
                textAlign: 'center',
              }}
            >
              {paragraph.bullet.char ?? '\u2022'}
            </span>
          ) : null}
          {paragraph.runs.map((run, runIndex) => (
            <TextRunRenderer
              key={runIndex}
              run={run}
              boxStyle={style}
              sourceId={`${element.id}:${paragraphIndex}:${runIndex}`}
              interactive={interactive}
            />
          ))}
        </div>
      ))}
    </div>
  );
}

export const TextRenderer = memo(TextRendererComponent);
