// TextRenderer 渲染 PPTX 文本框，并处理文本框形状、渐变填充、段落和 run 样式。
import type { CSSProperties } from 'react';
import React, { memo } from 'react';
import { resolveOfficeCssFontWeight } from '../../../services/fonts/OfficeFontResolver';
import type { OfficeFontFamilyResolver } from '../../../services/fonts/types';
import type {
  ReflectionStyle,
  TextElement,
  TextParagraph,
  TextRun,
  TextStyle,
} from '../../../services/pptx/types';
import { useOfficeFontResolver } from '../../../shared/fonts/OfficeFontProvider';
import { useOfficeHyperlink } from '../../../shared/hyperlink';
import { OfficeSearchHighlightedText } from '../../search/OfficeSearchContext';
import { reflectionCopyToCss, reflectionToCss } from './effects';
import { resolvePptxFontFamily } from './fontFamily';
import {
  resolvePresentationLeadingOffset,
  resolvePresentationLineHeight,
} from './textStyle';
import {
  colorWithOpacity,
  gradientToSvgEndpoints,
  isGradientPaint,
  paintToCss,
} from './paint';
import { buildRendererId } from './renderIds';
import { usePresentationTextAutoFit } from './usePresentationTextAutoFit';

/** Office 允许行末中文标点部分悬挂到文本边界之外。 */
const HANGING_END_PUNCTUATION = /[，。；：！？、）》】”’]$/;

/** 文本渲染器组件属性。 */
type TextRendererProps = {
  /** 当前处理或渲染的演示文稿元素。 */
  element: TextElement;
  /** 内容变化时用于刷新渲染结果的键。 */
  renderKey: string;
  /** 是否允许当前文本框及文字片段响应链接交互。 */
  interactive: boolean;
  /** 主视口中的零基幻灯片索引；未提供时不渲染查找高亮。 */
  searchSlideIndex?: number;
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

/** Office 会让普通文本字体中的符号继承文字颜色，避免浏览器改用不可着色的彩色 emoji。 */
function usesOfficeTextSymbols(fontFamily?: string) {
  return !fontFamily || !/emoji/i.test(fontFamily);
}

function equalReflection(
  left: ReflectionStyle,
  right: ReflectionStyle | undefined,
) {
  if (!right) return false;
  return (
    left.blur === right.blur &&
    left.distance === right.distance &&
    left.direction === right.direction &&
    left.startOpacity === right.startOpacity &&
    left.endOpacity === right.endOpacity &&
    left.startPosition === right.startPosition &&
    left.endPosition === right.endPosition
  );
}

/** 合并连续且倒影参数相同的文字片段，使显式换行仍在同一个倒影盒中生效。 */
function groupRunsByReflection(paragraph: TextParagraph) {
  return paragraph.runs.reduce<
    Array<{
      reflection?: ReflectionStyle;
      runs: Array<{ run: TextRun; index: number }>;
    }>
  >((groups, run, index) => {
    // PowerPoint 按视觉行生成文字倒影；显式换行必须切断分组，不能把多行整体翻转。
    if (run.text === '\n') {
      groups.push({ runs: [{ run, index }] });
      return groups;
    }
    const reflection = run.style?.reflection;
    const previous = groups[groups.length - 1];
    const sameReflection = reflection
      ? equalReflection(reflection, previous?.reflection)
      : !previous?.reflection;
    if (previous && sameReflection) {
      previous.runs.push({ run, index });
    } else {
      groups.push({ reflection, runs: [{ run, index }] });
    }
    return groups;
  }, []);
}

function isEmptyParagraph(paragraph: TextParagraph) {
  return paragraph.runs.every((run) => run.text.length === 0);
}

function resolveBulletLayout(paragraph: TextParagraph) {
  const hasBullet = Boolean(paragraph.bullet && !paragraph.bullet.none);
  const textIndent = paragraph.style?.textIndent ?? 0;
  const hangingWidth = hasBullet && textIndent < 0 ? -textIndent : undefined;
  return {
    paragraphPadding: hasBullet && !hangingWidth ? 18 : 0,
    bulletWidth: hangingWidth ? Math.max(12, hangingWidth - 6) : 12,
  };
}

/** 返回段落 strut 与未指定字号的项目符号共同使用的最小文字字号。 */
function resolveParagraphFontSize(
  paragraph: TextParagraph,
  boxStyle: TextStyle,
) {
  return (
    paragraph.runs.reduce<number | undefined>((minimum, run) => {
      if (!run.text.length) return minimum;
      const fontSize = run.style?.fontSize ?? boxStyle.fontSize;
      if (!fontSize) return minimum;
      return minimum === undefined ? fontSize : Math.min(minimum, fontSize);
    }, undefined) ?? boxStyle.fontSize
  );
}

function shouldHangEndPunctuation(
  paragraph: TextParagraph,
  runIndex: number,
) {
  for (let index = paragraph.runs.length - 1; index >= 0; index -= 1) {
    const text = paragraph.runs[index].text.trimEnd();
    if (!text) continue;
    return index === runIndex && HANGING_END_PUNCTUATION.test(text);
  }
  return false;
}

/** 渲染单个文字片段，使链接 Hook 不违反 React 循环调用约束。 */
function TextRunRenderer({
  run,
  boxStyle,
  sourceId,
  interactive,
  searchSlideIndex,
  elementId,
  resolveFontFamily,
  hangEndPunctuation,
  fontScale,
  trackingAdjustment,
}: {
  /** 当前文字片段。 */
  run: TextRun;
  /** 文本框继承的基础样式。 */
  boxStyle: TextStyle;
  /** 当前片段在文稿中的稳定来源标识。 */
  sourceId: string;
  /** 是否允许当前文字片段响应链接交互。 */
  interactive: boolean;
  /** 主视口中的零基幻灯片索引。 */
  searchSlideIndex?: number;
  /** 当前文本框的稳定标识。 */
  elementId: string;
  /** 当前文档会话统一的字体链解析函数。 */
  resolveFontFamily: OfficeFontFamilyResolver;
  /** 是否压缩段尾标点占位，模拟 Office 的标点悬挂。 */
  hangEndPunctuation?: boolean;
  /** 当前文本框自动适应后的字号缩放比例。 */
  fontScale: number;
  /** 用于补偿浏览器字体推进宽度差异的附加字距。 */
  trackingAdjustment: number;
}) {
  const hyperlinkProps = useOfficeHyperlink<HTMLSpanElement>({
    hyperlink: run.hyperlink,
    source: { type: 'text', id: sourceId },
    interactive,
  });
  const runStyle = run.style ?? {};
  const underline = runStyle.underline ?? boxStyle.underline;
  const strike = runStyle.strike ?? boxStyle.strike;
  const sourceFontFamily = runStyle.fontFamily ?? boxStyle.fontFamily;
  const sourceEastAsiaFontFamily =
    runStyle.eastAsiaFontFamily ?? boxStyle.eastAsiaFontFamily;
  const trimmedText = hangEndPunctuation ? run.text.trimEnd() : run.text;
  const hangingPunctuation =
    hangEndPunctuation && HANGING_END_PUNCTUATION.test(trimmedText)
      ? trimmedText.slice(-1)
      : undefined;
  const leadingText = hangingPunctuation
    ? trimmedText.slice(0, -1)
    : run.text;
  const trailingText = hangingPunctuation
    ? run.text.slice(trimmedText.length)
    : '';
  const renderText = (text: string) =>
    searchSlideIndex === undefined ? (
      text
    ) : (
      <OfficeSearchHighlightedText
        text={text}
        target={{
          kind: 'presentation',
          slideIndex: searchSlideIndex,
          elementId,
        }}
      />
    );
  return (
    <span
      {...hyperlinkProps}
      className={
        usesOfficeTextSymbols(sourceFontFamily)
          ? 'office-file-pptx-viewer__text-run--text-symbols'
          : undefined
      }
      style={{
        ...resolveTextPaintStyle(
          runStyle.textFill ?? boxStyle.textFill,
          runStyle.color ?? boxStyle.color ?? '#172033',
          runStyle.opacity ?? boxStyle.opacity,
        ),
        fontFamily: resolvePptxFontFamily(
          resolveFontFamily,
          sourceFontFamily,
          sourceEastAsiaFontFamily,
        ),
        fontSize:
          (runStyle.fontSize ?? boxStyle.fontSize) !== undefined
            ? (runStyle.fontSize ?? boxStyle.fontSize)! * fontScale
            : undefined,
        fontWeight: resolveOfficeCssFontWeight(
          runStyle.fontWeight ?? boxStyle.fontWeight,
          runStyle.bold ?? boxStyle.bold,
        ),
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
        letterSpacing:
          (runStyle.charSpace ?? boxStyle.charSpace ?? 0) * fontScale +
          trackingAdjustment,
      }}
    >
      {renderText(leadingText)}
      {hangingPunctuation ? (
        <span style={{ letterSpacing: '-0.75em' }}>
          {hangingPunctuation}
        </span>
      ) : null}
      {trailingText}
    </span>
  );
}

/** 渲染文本渲染器。 */
function TextRendererComponent({
  element,
  renderKey,
  interactive,
  searchSlideIndex,
}: TextRendererProps) {
  const resolveFontFamily = useOfficeFontResolver();
  const style = element.boxStyle ?? {};
  const hasVisibleShapeBoundary = Boolean(
    (element.fill !== undefined &&
      element.fill !== null &&
      element.fillOpacity !== 0) ||
      (element.stroke && element.strokeOpacity !== 0) ||
      element.shadow,
  );
  const {
    containerRef,
    scale: autoFitScale,
    tracking: autoFitTracking,
  } = usePresentationTextAutoFit(
    style.fit,
    searchSlideIndex !== undefined &&
      (style.fit === 'shrinkText' ||
        (style.fit === 'resizeShape' && hasVisibleShapeBoundary)),
    `${renderKey}:${element.id}:${element.width}:${element.height}`,
  );
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
  const hasRenderableText = element.paragraphs.some(
    (paragraph) =>
      paragraph.runs.some((run) => run.text.length > 0) ||
      Boolean(paragraph.bullet && !paragraph.bullet.none),
  );
  const renderedParagraphs = element.paragraphs.map((paragraph, index) => ({
    paragraph,
    index,
  }));
  // OOXML 文本框常带一个仅用于保存结束格式的空末段；Office 不把它计入可见排版。
  if (
    hasRenderableText &&
    renderedParagraphs.length > 1 &&
    isEmptyParagraph(renderedParagraphs[renderedParagraphs.length - 1].paragraph)
  ) {
    renderedParagraphs.pop();
  }
  const strokeWidth = element.strokeWidth ?? 1;
  const isHorizontalRule = Boolean(
    !isVectorShape && element.stroke && element.height <= strokeWidth,
  );
  const isVerticalRule = Boolean(
    !isVectorShape && element.stroke && element.width <= strokeWidth,
  );
  const strokeCss = element.stroke
    ? `${strokeWidth}px ${lineStyle(element.strokeDash)} ${colorWithOpacity(
        element.stroke,
        element.strokeOpacity,
      )}`
    : undefined;
  const shapeBorderStyle: CSSProperties = isHorizontalRule
    ? { borderTop: strokeCss }
    : isVerticalRule
    ? { borderLeft: strokeCss }
    : !isVectorShape && strokeCss
    ? { border: strokeCss }
    : {};
  // 同一页可能在缩略图和主视口同时出现，SVG defs id 需要带 renderKey 防止互相引用错。
  const gradientId = isGradientFill
    ? buildRendererId(renderKey, element.id, 'fill-gradient')
    : undefined;
  const verticalJustify =
    style.verticalAlign === 'bottom'
      ? 'flex-end'
      : style.verticalAlign === 'middle'
      ? 'center'
      : 'flex-start';
  const hasTextReflection = element.paragraphs.some((paragraph) =>
    paragraph.runs.some((run) => Boolean(run.style?.reflection)),
  );
  const paragraphCss = (paragraph: TextParagraph): CSSProperties => {
    // 段落 strut 使用最小字号，让混合字号软换行后的每一行由实际 run 决定高度。
    const paragraphFontSize = resolveParagraphFontSize(paragraph, style);
    const scaledParagraphFontSize =
      paragraphFontSize !== undefined
        ? paragraphFontSize * autoFitScale
        : undefined;
    const lineHeight = (() => {
      const value = resolvePresentationLineHeight(
        paragraph.style?.lineHeight ?? style.lineHeight,
      );
      return typeof value === 'string'
        ? `${Number.parseFloat(value) * autoFitScale}px`
        : value;
    })();
    const bulletLayout = resolveBulletLayout(paragraph);
    return {
      position: 'relative',
      // 仅可见形状需要补偿固定行距的首行留白；透明标题框按 PowerPoint 原始坐标起排。
      top:
        style.fit === 'resizeShape' && hasVisibleShapeBoundary
          ? resolvePresentationLeadingOffset(
              scaledParagraphFontSize,
              lineHeight,
            )
          : undefined,
      zIndex: 1,
      textAlign: paragraph.style?.align ?? style.align ?? 'left',
      lineHeight,
      fontSize: scaledParagraphFontSize,
      marginTop: (paragraph.style?.spaceBefore ?? 0) * autoFitScale,
      marginBottom: (paragraph.style?.spaceAfter ?? 0) * autoFitScale,
      paddingLeft: `${
        (paragraph.style?.marginLeft ?? 0) +
        bulletLayout.paragraphPadding
      }px`,
      textIndent: paragraph.style?.textIndent
        ? `${paragraph.style.textIndent}px`
        : undefined,
      display: 'block',
      whiteSpace: 'inherit',
    };
  };

  return (
    <div
      ref={containerRef}
      {...hyperlinkProps}
      data-office-presentation-element-id={
        searchSlideIndex === undefined ? undefined : element.id
      }
      data-office-presentation-slide-index={searchSlideIndex}
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
        fontFamily: resolvePptxFontFamily(
          resolveFontFamily,
          style.fontFamily,
          style.eastAsiaFontFamily,
        ),
        fontSize:
          style.fontSize !== undefined
            ? style.fontSize * autoFitScale
            : undefined,
        fontWeight: resolveOfficeCssFontWeight(style.fontWeight, style.bold),
        fontStyle: style.italic ? 'italic' : 'normal',
        textDecoration: textDecoration(style),
        textTransform: style.allCaps ? 'uppercase' : undefined,
        fontVariant: style.smallCaps ? 'small-caps' : undefined,
        writingMode: style.writingMode,
        whiteSpace: 'pre-wrap',
        lineHeight: (() => {
          const lineHeight = resolvePresentationLineHeight(style.lineHeight);
          return typeof lineHeight === 'string'
            ? `${Number.parseFloat(lineHeight) * autoFitScale}px`
            : lineHeight;
        })(),
        background: isVectorShape
          ? undefined
          : paintToCss(element.fill, element.fillOpacity),
        ...shapeBorderStyle,
        borderRadius: radiusToPx(element),
        boxShadow: shadowToCss(element),
        WebkitBoxReflect: reflectionToCss(element.reflection),
        transform: element.rotate
          ? `rotate(${element.rotate}deg)${element.flipH ? ' scaleX(-1)' : ''}${
              element.flipV ? ' scaleY(-1)' : ''
            }`
          : `${element.flipH ? 'scaleX(-1)' : ''}${
              element.flipV ? ' scaleY(-1)' : ''
            }`.trim() || undefined,
        transformOrigin: 'center center',
        boxSizing: 'border-box',
        // 把段落和反射的 z-index 限制在当前形状内，避免文字越过后续形状破坏 PPT 绘制顺序。
        isolation: 'isolate',
        // PowerPoint 即使关闭自动调整也会绘制越过文本框边界的文字，不能按 HTML 容器默认裁切。
        overflow: 'visible',
        display: 'flex',
        flexDirection: 'column',
        justifyContent: verticalJustify,
        paddingLeft: hasRenderableText ? style.marginLeft ?? 0 : 0,
        paddingRight: hasRenderableText ? style.marginRight ?? 0 : 0,
        paddingTop: hasRenderableText ? style.marginTop ?? 0 : 0,
        paddingBottom: hasRenderableText ? style.marginBottom ?? 0 : 0,
        letterSpacing:
          (style.charSpace ?? 0) * autoFitScale + autoFitTracking,
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
      {renderedParagraphs.map(({ paragraph, index: paragraphIndex }) => (
        <div
          key={paragraphIndex}
          data-office-presentation-text-paragraph="true"
          style={paragraphCss(paragraph)}
        >
          {paragraph.bullet && !paragraph.bullet.none ? (
            <span
              style={{
                display: 'inline-block',
                color: colorWithOpacity(
                  paragraph.bullet.color ?? style.color,
                  style.opacity,
                ),
                fontSize:
                  (paragraph.bullet.size ??
                    resolveParagraphFontSize(paragraph, style)) !== undefined
                    ? (paragraph.bullet.size ??
                        resolveParagraphFontSize(paragraph, style))! *
                      autoFitScale
                    : undefined,
                fontFamily: resolveFontFamily(
                  paragraph.bullet.fontFamily ?? style.fontFamily,
                ),
                marginRight: 6,
                width: resolveBulletLayout(paragraph).bulletWidth,
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
              searchSlideIndex={searchSlideIndex}
              elementId={element.id}
              resolveFontFamily={resolveFontFamily}
              hangEndPunctuation={shouldHangEndPunctuation(
                paragraph,
                runIndex,
              )}
              fontScale={autoFitScale}
              trackingAdjustment={autoFitTracking}
            />
          ))}
          {hasRenderableText && isEmptyParagraph(paragraph) ? (
            <span aria-hidden="true" style={{ visibility: 'hidden' }}>
              {'\u00a0'}
            </span>
          ) : null}
        </div>
      ))}
      {hasTextReflection ? (
        <div
          aria-hidden="true"
          style={{
            position: 'absolute',
            left: style.marginLeft ?? 0,
            right: style.marginRight ?? 0,
            top: style.marginTop ?? 0,
            bottom: style.marginBottom ?? 0,
            zIndex: 2,
            display: 'flex',
            flexDirection: 'column',
            justifyContent: verticalJustify,
            pointerEvents: 'none',
            whiteSpace: 'pre-wrap',
            overflow: 'visible',
          }}
        >
          {renderedParagraphs.map(({ paragraph, index: paragraphIndex }) => (
            <div
              key={paragraphIndex}
              style={{ ...paragraphCss(paragraph), zIndex: 2 }}
            >
              {paragraph.bullet && !paragraph.bullet.none ? (
                <span
                  style={{
                    display: 'inline-block',
                    visibility: 'hidden',
                    fontSize:
                      (paragraph.bullet.size ??
                        resolveParagraphFontSize(paragraph, style)) !== undefined
                        ? (paragraph.bullet.size ??
                            resolveParagraphFontSize(paragraph, style))! *
                          autoFitScale
                        : undefined,
                    fontFamily: resolveFontFamily(
                      paragraph.bullet.fontFamily ?? style.fontFamily,
                    ),
                    marginRight: 6,
                    width: resolveBulletLayout(paragraph).bulletWidth,
                    textAlign: 'center',
                  }}
                >
                  {paragraph.bullet.char ?? '\u2022'}
                </span>
              ) : null}
              {groupRunsByReflection(paragraph).map((group, groupIndex) => {
                const hiddenContent = group.runs.map(({ run, index }) => (
                  <TextRunRenderer
                    key={`layout-${index}`}
                    run={run}
                    boxStyle={style}
                    sourceId={`${element.id}:${paragraphIndex}:${index}:layout`}
                    interactive={false}
                    elementId={element.id}
                    resolveFontFamily={resolveFontFamily}
                    hangEndPunctuation={shouldHangEndPunctuation(
                      paragraph,
                      index,
                    )}
                    fontScale={autoFitScale}
                    trackingAdjustment={autoFitTracking}
                  />
                ));
                if (!group.reflection) {
                  return (
                    <span key={groupIndex} style={{ visibility: 'hidden' }}>
                      {hiddenContent}
                    </span>
                  );
                }
                return (
                  <span
                    key={groupIndex}
                    data-office-text-reflection="true"
                    style={{
                      position: 'relative',
                      display: 'inline-block',
                      maxWidth: '100%',
                      verticalAlign: 'top',
                      whiteSpace: 'inherit',
                    }}
                  >
                    <span style={{ visibility: 'hidden' }}>
                      {hiddenContent}
                    </span>
                    <span
                      data-office-text-reflection-copy="true"
                      style={{
                        display: 'inline-block',
                        whiteSpace: 'inherit',
                        ...reflectionCopyToCss(group.reflection),
                      }}
                    >
                      {group.runs.map(({ run, index }) => (
                        <TextRunRenderer
                          key={`reflection-${index}`}
                          run={run}
                          boxStyle={style}
                          sourceId={`${element.id}:${paragraphIndex}:${index}:reflection`}
                          interactive={false}
                          elementId={element.id}
                          resolveFontFamily={resolveFontFamily}
                          hangEndPunctuation={shouldHangEndPunctuation(
                            paragraph,
                            index,
                          )}
                          fontScale={autoFitScale}
                          trackingAdjustment={autoFitTracking}
                        />
                      ))}
                    </span>
                  </span>
                );
              })}
              {hasRenderableText && isEmptyParagraph(paragraph) ? (
                <span aria-hidden="true" style={{ visibility: 'hidden' }}>
                  {'\u00a0'}
                </span>
              ) : null}
            </div>
          ))}
        </div>
      ) : null}
    </div>
  );
}

export const TextRenderer = memo(TextRendererComponent);
