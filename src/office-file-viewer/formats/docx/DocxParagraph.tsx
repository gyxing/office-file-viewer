// DocxParagraph 渲染 DOCX 段落块，并应用段落级缩进、间距、边框和文字样式。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { DocxParagraphBlock } from '../../services/docx/types';
import { useOfficeFontResolver } from '../../shared/fonts/OfficeFontProvider';
import { DocxInlineContent } from './DocxInlineContent';
import {
  buildDocxTextStyle,
  getDocxCssLineHeight,
  getDocxEmptyParagraphHeight,
  resolveDocxShapeBaselineOffset,
  resolveDocxTextFontFamily,
} from './docxRenderUtils';
import { calculatePositionStyle } from './positionUtils';

/** DOCX段落组件属性。 */
type DocxParagraphProps = {
  /** 当前负责处理或渲染的内容块。 */
  block: DocxParagraphBlock;
  /** 是否使用紧凑段落布局。 */
  compact?: boolean;
  /** 是否使用 div 元素承载当前段落。 */
  asDiv?: boolean; // 强制使用 div 而不是 p,用于避免嵌套问题
  /** 是否使用 Office 形状文本的字体场景基线。 */
  insideShape?: boolean;
  /** 查找结果对应的顶层正文块标识。 */
  searchBlockId?: string;
};

/** 渲染DOCX段落。 */
function DocxParagraphComponent({
  block,
  compact = false,
  asDiv = false,
  insideShape = false,
  searchBlockId = block.sourceBlockId ?? block.id,
}: DocxParagraphProps) {
  const resolveFontFamily = useOfficeFontResolver();
  // 浏览器段落 strut 需采用首个实际文字运行的字体指标，否则不同字体基线会额外撑高每一行。
  const lineMetricInline = block.inlines.find(
    (inline) => inline.type === 'text' && Boolean(inline.text),
  );
  const lineMetricStyle =
    lineMetricInline?.type === 'text' ? lineMetricInline.style : block.style;
  const shapeBaselineOffset = insideShape
    ? resolveDocxShapeBaselineOffset(lineMetricStyle)
    : undefined;
  const hasContent = block.inlines.length > 0;
  const hasFlowContent = block.inlines.some((inline) => {
    if (inline.type === 'text') return inline.text.length > 0;
    if (inline.type === 'tab') return true;
    if (inline.type === 'break') return true;
    if (inline.type === 'image') return !inline.image.position;
    if (inline.type === 'shape') return !inline.shape.position;
    if (inline.type === 'chart') return !inline.chart.position;
    return false;
  });

  // 检查是否包含定位元素,如果包含则使用 div 而不是 p 以避免 DOM 嵌套警告
  const hasPositionedElements = block.inlines.some((inline) => {
    if (inline.type === 'image') return Boolean(inline.image.position);
    if (inline.type === 'shape') return Boolean(inline.shape.position);
    if (inline.type === 'chart') return Boolean(inline.chart.position);
    return false;
  });
  const hasBlockLevelInline = block.inlines.some(
    (inline) => inline.type === 'shape' || inline.type === 'chart',
  );
  const tocTabIndex = block.isTableOfContents
    ? block.inlines.findIndex((inline) => inline.type === 'tab')
    : -1;
  const isTocLine = tocTabIndex >= 0;

  const positionStyle = calculatePositionStyle(block.position);

  const paragraphStyle = useMemo<CSSProperties>(() => {
    // Word 会保留浮动对象所在的锚点行；若压成 0 高度，后续按段落定位的对象会发生累计偏移。
    const lineMetricCss = buildDocxTextStyle(
      lineMetricStyle,
      undefined,
      resolveFontFamily,
    );
    const baseMinHeight = hasFlowContent
      ? undefined
      : getDocxEmptyParagraphHeight(block);
    return {
      ...positionStyle,
      position: block.position ? positionStyle.position : 'relative',
      transform:
        !block.position && shapeBaselineOffset !== undefined
          ? `translateY(${shapeBaselineOffset}px)`
          : positionStyle.transform,
      zIndex: block.position
        ? positionStyle.zIndex
        : hasFlowContent
        ? 1
        : undefined,
      margin: block.position ? 0 : undefined,
      marginTop: block.position
        ? undefined
        : compact
        ? 0
        : block.spacingBefore ?? 0,
      marginRight: block.position ? undefined : block.indentRight,
      marginBottom: block.position ? undefined : block.spacingAfter ?? 0,
      marginLeft: block.position ? undefined : block.indentLeft,
      paddingLeft: block.paddingLeft,
      paddingRight: block.paddingRight,
      minHeight: baseMinHeight,
      textAlign: block.align,
      // flex 基线布局会把绝对行距向上取整约 1px，目录行需抵消该误差。
      lineHeight:
        isTocLine && block.lineHeight !== undefined && block.lineHeight > 4
          ? `${Math.max(1, block.lineHeight - 1)}px`
          : getDocxCssLineHeight(block),
      color: block.style?.color ?? '#000',
      background: block.backgroundColor,
      borderTop: block.borderTop,
      borderRight: block.borderRight,
      borderBottom: block.borderBottom,
      borderLeft: block.borderLeft,
      textIndent: block.firstLineIndent,
      paddingTop: block.paddingTop,
      paddingBottom: block.paddingBottom,
      maxWidth: block.position ? 'none' : undefined,
      ...buildDocxTextStyle(block.style, undefined, resolveFontFamily),
      fontSize: lineMetricStyle?.fontSize ?? block.style?.fontSize ?? 14,
      fontFamily:
        resolveDocxTextFontFamily(lineMetricStyle, resolveFontFamily) ??
        resolveDocxTextFontFamily(block.style, resolveFontFamily),
      fontWeight: lineMetricCss.fontWeight,
      WebkitTextStroke: lineMetricCss.WebkitTextStroke,
    };
  }, [
    block,
    compact,
    shapeBaselineOffset,
    hasContent,
    hasFlowContent,
    isTocLine,
    lineMetricStyle,
    positionStyle,
    resolveFontFamily,
  ]);

  // 图表和形状内部会渲染块级节点，使用 div 容器避免嵌套到 p 里触发浏览器修正。
  const Container =
    hasPositionedElements || hasBlockLevelInline || asDiv ? 'div' : 'p';
  const paragraphEndInlineIndex = block.inlines.reduce(
    (lastIndex, inline, index) => (inline.type === 'text' ? index : lastIndex),
    -1,
  );
  const tocLeader = block.tabStops?.find(
    (stop) => stop.align === 'right' || stop.align === 'decimal',
  )?.leader;
  const renderInline = (
    inline: (typeof block.inlines)[number],
    index: number,
  ) => (
    <DocxInlineContent
      key={`${block.id}-inline-${index}`}
      inline={inline}
      sourceId={`${block.id}-inline-${index}`}
      searchBlockId={searchBlockId}
      resolveFontFamily={resolveFontFamily}
      isParagraphEnd={index === paragraphEndInlineIndex}
    />
  );

  return (
    <Container
      className={isTocLine ? 'office-file-docx-toc-line' : undefined}
      style={paragraphStyle}
      data-office-word-outline-target={
        block.outlineLevel !== undefined ? block.id : undefined
      }
      data-office-word-block-id={searchBlockId}
      data-office-docx-align={block.align}
    >
      {tocTabIndex >= 0 ? (
        <>
          <span className="office-file-docx-toc-line__text">
            {block.inlines.slice(0, tocTabIndex).map(renderInline)}
          </span>
          <span
            className="office-file-docx-toc-line__leader"
            data-office-docx-tab-leader={tocLeader ?? 'none'}
            aria-hidden="true"
          />
          <span className="office-file-docx-toc-line__page">
            {block.inlines
              .slice(tocTabIndex + 1)
              .map((inline, index) =>
                renderInline(inline, tocTabIndex + index + 1),
              )}
          </span>
        </>
      ) : (
        block.inlines.map(renderInline)
      )}
    </Container>
  );
}

export const DocxParagraph = memo(DocxParagraphComponent);
