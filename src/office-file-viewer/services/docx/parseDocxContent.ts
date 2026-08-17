import type { OfficeHyperlink } from '../../shared/hyperlink';
import { readXml } from '../../shared/ooxml/archive';
import type { OfficeTheme } from '../../shared/ooxml/theme';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  descendantsByLocalName,
  matchesLocalName,
  parseXml,
  textContent,
} from '../../shared/ooxml/xml';
import {
  DEFAULT_DOCX_PAGE,
  mapAlignment,
  mergeTextStyle,
  positiveTwipToPx,
  readBorder,
  readDocxTableBorders,
  readOnOff,
  readShading,
  readVal,
  resolveDocxStyle,
  resolveDocxTableBorders,
  resolveDocxTableCellMargins,
  resolveParagraphStyle,
  resolveRunStyle,
  twipToPx,
  type DocxParseContext,
  type ReadBlockChildrenOptions,
} from './docxParsingContext';
import type { DocxBlockParseOperations } from './parseDocxBlock';
import { createDocxDrawingParser } from './parseDocxDrawing';
import {
  parseDocxFieldHyperlink,
  parseDocxHyperlinkElement,
} from './parseDocxHyperlink';
import { nextDocxNumberPrefix } from './parseDocxNumbering';
import type {
  DocxBlock,
  DocxInline,
  DocxPage,
  DocxPageContent,
  DocxPageRegionVariants,
  DocxParagraphBlock,
  DocxPosition,
  DocxTableBlock,
  DocxTableCell,
  DocxTableRow,
  DocxTextStyle,
} from './types';

/** DOCX 内容解析流程共享的样式、关系和资源上下文。 */
type ParseContext = DocxParseContext;

const drawingParser = createDocxDrawingParser(readDocxBlockChildren);

// 百分比表宽不包含首尾默认单元格边距，Word 会把两侧各约 7px 绘制到正文边界外。
/** DOCX 表格缺少边缘设置时使用的默认修正量。 */
const DOCX_DEFAULT_TABLE_EDGE_OFFSET = 7;
/** Word/WPS 正文单倍行距的默认行盒倍率，用于判断段落占用的文档网格数。 */
const DOCX_BODY_LINE_HEIGHT_MULTIPLIER = 4 / 3;
/** WPS 非表格 auto 行距使用的 12 磅排版基准。 */
const DOCX_AUTO_LINE_HEIGHT_BASE_FONT_SIZE_PX = 16;
/** Word 未显式声明行距的标题使用的自然行盒倍率。 */
const DOCX_HEADING_LINE_HEIGHT_MULTIPLIER = 1.8;
/** WPS 纯浮动锚点段落使用微软雅黑段落标记时采用的兼容行盒倍率。 */
const DOCX_ANCHOR_YAHEI_LINE_HEIGHT_MULTIPLIER = 1.8;
/** 可按字体场景还原纯浮动锚点段落网格占用的微软雅黑字体。 */
const DOCX_YAHEI_FONT_PATTERN = /微软雅黑|Microsoft\s+YaHei/i;
/** 宋体类字体在 12pt 下可完整落入 15.6pt 单行网格，需避免误吸附为双行网格。 */
const DOCX_COMPACT_SERIF_LINE_HEIGHT_MULTIPLIER = 1.3;
/** Word/WPS 文本框兼容行距的默认行盒倍率，用于还原文本框内的网格吸附。 */
const DOCX_SHAPE_LINE_HEIGHT_MULTIPLIER = 1.2;
/** 使用紧凑东亚衬线字形度量的常见 Office 字体。 */
const DOCX_COMPACT_SERIF_FONT_PATTERN =
  /宋体|新宋体|ＭＳ 明朝|SimSun|NSimSun|MS Mincho|Songti|Noto Serif CJK/i;
/** 编号标签挤满悬挂区后，Word 在正文前保留的最小视觉间隔。 */
const DOCX_NUMBERING_OVERFLOW_GAP_EM = 0.875;

/** 估算编号标签占用宽度，避免长编号越过悬挂区后与正文重叠。 */
function estimateDocxNumberingLabelWidth(text: string, fontSize: number) {
  return Array.from(text).reduce((width, character) => {
    if (/[0-9]/.test(character)) return width + fontSize * 0.5;
    if (/[A-Za-z]/.test(character)) return width + fontSize * 0.55;
    if (/[.,:;]/.test(character)) return width + fontSize * 0.25;
    if (/\s/.test(character)) return width + fontSize * 0.33;
    return width + fontSize;
  }, 0);
}

/** 计算编号文字溢出悬挂区后，Word 在正文前追加的间隔。 */
function resolveDocxNumberingGapWidth(
  text: string,
  hangingWidth: number,
  fontSize: number,
) {
  const labelWidth = estimateDocxNumberingLabelWidth(text, fontSize);
  return labelWidth >= hangingWidth
    ? labelWidth - hangingWidth + fontSize * DOCX_NUMBERING_OVERFLOW_GAP_EM
    : 0;
}

/** 复杂域解析时保留的指令和已确定链接。 */
type DocxFieldState = {
  /** 当前字段累计读取的指令文本。 */
  instruction: string;
  /** 字段进入结果区后可绑定到可见内容的链接。 */
  hyperlink?: OfficeHyperlink;
};

function applyHyperlinkToInline(
  inline: DocxInline,
  hyperlink: OfficeHyperlink | undefined,
): DocxInline {
  if (!hyperlink) return inline;
  if (inline.type === 'text') return { ...inline, hyperlink };
  if (inline.type === 'image') {
    return { ...inline, image: { ...inline.image, hyperlink } };
  }
  if (inline.type === 'shape') {
    return { ...inline, shape: { ...inline.shape, hyperlink } };
  }
  return inline;
}

function currentFieldHyperlink(fieldStack: readonly DocxFieldState[]) {
  for (let index = fieldStack.length - 1; index >= 0; index -= 1) {
    if (fieldStack[index].hyperlink) return fieldStack[index].hyperlink;
  }
  return undefined;
}

/** 形状文字只声明脚本提示时，Word 会改用 DrawingML 次要主题字体度量西文。 */
function resolveShapeThemeFontFamily(
  rPr: Element | null | undefined,
  context: ParseContext,
  insideShape: boolean,
) {
  if (!insideShape) return undefined;
  const fonts = childByLocalName(rPr, 'rFonts');
  if (!fonts) return undefined;
  const explicitFontAttributes = [
    'ascii',
    'hAnsi',
    'eastAsia',
    'cs',
    'asciiTheme',
    'hAnsiTheme',
    'eastAsiaTheme',
    'cstheme',
  ];
  const hasExplicitFont = explicitFontAttributes.some(
    (name) => attr(fonts, `w:${name}`) ?? attr(fonts, name),
  );
  if (hasExplicitFont) return undefined;
  return context.theme.fontScheme?.minorFont;
}

/** 判断两个文字运行是否具有完全一致的可见样式。 */
function haveEqualTextStyles(
  left: DocxTextStyle | undefined,
  right: DocxTextStyle | undefined,
) {
  const leftEntries = Object.entries(left ?? {});
  const rightEntries = Object.entries(right ?? {});
  return (
    leftEntries.length === rightEntries.length &&
    leftEntries.every(
      ([name, value]) => right?.[name as keyof DocxTextStyle] === value,
    )
  );
}

/** 合并跨 run 的同样式连续空格，避免浏览器在相邻 span 边界错误压缩列距。 */
function mergeAdjacentPreservedSpaces(inlines: DocxInline[]) {
  return inlines.reduce<DocxInline[]>((merged, inline) => {
    const previous = merged[merged.length - 1];
    if (
      previous?.type === 'text' &&
      inline.type === 'text' &&
      previous.text.endsWith(' ') &&
      inline.text.startsWith(' ') &&
      previous.hyperlink === inline.hyperlink &&
      haveEqualTextStyles(previous.style, inline.style)
    ) {
      previous.text += inline.text;
      previous.preserveSpace = true;
      return merged;
    }
    merged.push(inline);
    return merged;
  }, []);
}

/** 将单个 run 的文本、换行和绘图子节点转换为行内模型。 */
function parseRun(
  runNode: Element,
  paragraphStyle: DocxTextStyle | undefined,
  context: ParseContext,
  fieldStack: DocxFieldState[],
  wrapperHyperlink?: OfficeHyperlink,
  insideShape = false,
): DocxInline[] {
  const rPr = childByLocalName(runNode, 'rPr');
  const shapeThemeFontFamily = resolveShapeThemeFontFamily(
    rPr,
    context,
    insideShape,
  );
  const runStyle = mergeTextStyle(
    resolveRunStyle(rPr, context.styles, context.theme, paragraphStyle),
    shapeThemeFontFamily ? { fontFamily: shapeThemeFontFamily } : undefined,
  );
  const inlines: DocxInline[] = [];

  Array.from(runNode.children).forEach((child) => {
    if (matchesLocalName(child, 'fldChar')) {
      const fieldType =
        attr(child, 'w:fldCharType') ?? attr(child, 'fldCharType');
      if (fieldType === 'begin') fieldStack.push({ instruction: '' });
      if (fieldType === 'separate' && fieldStack.length) {
        const field = fieldStack[fieldStack.length - 1];
        field.hyperlink = parseDocxFieldHyperlink(field.instruction);
      }
      if (fieldType === 'end') fieldStack.pop();
      return;
    }
    if (matchesLocalName(child, 'instrText')) {
      const field = fieldStack[fieldStack.length - 1];
      if (field) field.instruction += textContent(child);
      return;
    }
    const activeHyperlink =
      wrapperHyperlink ?? currentFieldHyperlink(fieldStack);
    if (matchesLocalName(child, 't')) {
      const text = textContent(child);
      inlines.push({
        type: 'text',
        text,
        // OOXML 会保留文本节点内部的连续空格；xml:space 额外覆盖首尾空格。
        preserveSpace:
          attr(child, 'xml:space') === 'preserve' || / {2,}/.test(text)
            ? true
            : undefined,
        style: runStyle,
        hyperlink: activeHyperlink,
      });
      return;
    }
    if (matchesLocalName(child, 'tab')) {
      inlines.push({ type: 'tab', style: runStyle });
      return;
    }
    if (matchesLocalName(child, 'br') || matchesLocalName(child, 'cr')) {
      inlines.push({ type: 'break' });
      return;
    }
    const drawingInline = drawingParser.parseRunChild(child, context);
    if (drawingInline) {
      inlines.push(applyHyperlinkToInline(drawingInline, activeHyperlink));
    }
  });

  return inlines;
}

/** 判断修订是否连同段落标记一起删除。 */
function hasDeletedParagraphMark(pNode: Element) {
  const paragraphProperties = childByLocalName(pNode, 'pPr');
  const paragraphMarkProperties = childByLocalName(paragraphProperties, 'rPr');
  return Array.from(paragraphMarkProperties?.children ?? []).some(
    (child) =>
      matchesLocalName(child, 'del') || matchesLocalName(child, 'moveFrom'),
  );
}

/** 判断段落在忽略删除修订后是否仍有需要保留的行内内容。 */
function hasFinalParagraphContent(paragraph: DocxParagraphBlock) {
  return paragraph.inlines.some((inline) => {
    if (inline.type === 'bookmark') return false;
    if (inline.type === 'text') return inline.text.length > 0;
    return true;
  });
}

function readParagraphBlocks(
  pNode: Element,
  id: string,
  context: ParseContext,
  options?: ReadBlockChildrenOptions,
): DocxParagraphBlock[] {
  const paragraph = parseParagraph(pNode, id, context, options);
  // 删除段落标记可能仅用于合并相邻段落；只有最终状态确实为空时才移除行盒。
  if (hasDeletedParagraphMark(pNode) && !hasFinalParagraphContent(paragraph)) {
    return [];
  }
  return [paragraph];
}

/** 按“最终状态”读取段落行内内容：保留插入/移入内容，忽略删除/移出内容。 */
function readParagraphRunChildren(
  parentNode: Element,
  paragraphStyle: DocxTextStyle | undefined,
  context: ParseContext,
  blockId: string,
  fieldStack: DocxFieldState[],
  wrapperHyperlink?: OfficeHyperlink,
  insideShape = false,
) {
  const inlines: DocxInline[] = [];

  Array.from(parentNode.children).forEach((child) => {
    if (matchesLocalName(child, 'r')) {
      inlines.push(
        ...parseRun(
          child,
          paragraphStyle,
          context,
          fieldStack,
          wrapperHyperlink,
          insideShape,
        ),
      );
      return;
    }
    if (matchesLocalName(child, 'bookmarkStart')) {
      const name = attr(child, 'w:name') ?? attr(child, 'name');
      if (!name) return;
      const markerId = `word-bookmark-${name}`;
      context.bookmarks[name] = { name, targetBlockId: blockId, markerId };
      inlines.push({ type: 'bookmark', name, markerId });
      return;
    }
    if (matchesLocalName(child, 'del') || matchesLocalName(child, 'moveFrom')) {
      return;
    }
    if (matchesLocalName(child, 'sdt')) {
      const content = childByLocalName(child, 'sdtContent');
      if (content) {
        inlines.push(
          ...readParagraphRunChildren(
            content,
            paragraphStyle,
            context,
            blockId,
            fieldStack,
            wrapperHyperlink,
            insideShape,
          ),
        );
      }
      return;
    }
    if (matchesLocalName(child, 'hyperlink')) {
      const hyperlink = parseDocxHyperlinkElement(child, context.documentRels);
      inlines.push(
        ...readParagraphRunChildren(
          child,
          paragraphStyle,
          context,
          blockId,
          fieldStack,
          hyperlink ?? wrapperHyperlink,
          insideShape,
        ),
      );
      return;
    }
    if (matchesLocalName(child, 'fldSimple')) {
      const instruction = attr(child, 'w:instr') ?? attr(child, 'instr') ?? '';
      const hyperlink = parseDocxFieldHyperlink(instruction);
      inlines.push(
        ...readParagraphRunChildren(
          child,
          paragraphStyle,
          context,
          blockId,
          fieldStack,
          hyperlink ?? wrapperHyperlink,
          insideShape,
        ),
      );
      return;
    }
    if (
      matchesLocalName(child, 'ins') ||
      matchesLocalName(child, 'moveTo') ||
      matchesLocalName(child, 'smartTag') ||
      matchesLocalName(child, 'customXml') ||
      matchesLocalName(child, 'sdtContent')
    ) {
      inlines.push(
        ...readParagraphRunChildren(
          child,
          paragraphStyle,
          context,
          blockId,
          fieldStack,
          wrapperHyperlink,
          insideShape,
        ),
      );
    }
  });

  return inlines;
}

function readParagraphRuns(
  pNode: Element,
  paragraphStyle: DocxTextStyle | undefined,
  context: ParseContext,
  blockId: string,
  insideShape = false,
) {
  return mergeAdjacentPreservedSpaces(
    readParagraphRunChildren(
      pNode,
      paragraphStyle,
      context,
      blockId,
      [],
      undefined,
      insideShape,
    ),
  );
}

function textFromInlines(inlines: DocxInline[]) {
  return inlines
    .map((inline) =>
      inline.type === 'text' ? inline.text : inline.type === 'tab' ? '\t' : '',
    )
    .join('');
}

/** 判断段落是否仅承载浮动对象及其空白锚点内容。 */
function containsOnlyPositionedInlines(inlines: DocxInline[]) {
  if (!inlines.length) return false;
  let containsPositionedObject = false;
  const onlyPositionedContent = inlines.every((inline) => {
    if (inline.type === 'text') return !inline.text.trim();
    if (inline.type === 'bookmark') return true;
    if (inline.type === 'break' || inline.type === 'tab') return false;
    if (inline.type === 'image') {
      containsPositionedObject ||= Boolean(inline.image.position);
      return Boolean(inline.image.position);
    }
    if (inline.type === 'shape') {
      containsPositionedObject ||= Boolean(inline.shape.position);
      return Boolean(inline.shape.position);
    }
    if (inline.type === 'chart') {
      containsPositionedObject ||= Boolean(inline.chart.position);
      return Boolean(inline.chart.position);
    }
    return false;
  });
  return onlyPositionedContent && containsPositionedObject;
}

function parseParagraph(
  pNode: Element,
  id: string,
  context: ParseContext,
  options?: ReadBlockChildrenOptions,
): DocxParagraphBlock {
  const pPr = childByLocalName(pNode, 'pPr');
  const style = resolveParagraphStyle(
    pPr,
    context.styles,
    context.theme,
    options?.paragraphContextStyle,
  );
  const inlines = readParagraphRuns(
    pNode,
    style.style,
    context,
    id,
    Boolean(options?.insideShape),
  );
  const sourceText = textFromInlines(inlines).trim();
  const numberingReference = style.numbering
    ? {
        ...style.numbering,
        level: style.numbering.level ?? style.outlineLevel ?? 0,
      }
    : undefined;
  const numberPrefix =
    sourceText && !options?.insidePageRegion && numberingReference
      ? nextDocxNumberPrefix(numberingReference, context.numbering)
      : undefined;
  const directIndent = childByLocalName(pPr, 'ind');
  const directIndentLeft = twipToPx(
    attr(directIndent, 'w:start') ??
      attr(directIndent, 'start') ??
      attr(directIndent, 'w:left') ??
      attr(directIndent, 'left'),
  );
  const directFirstLineIndent = twipToPx(
    attr(directIndent, 'w:firstLine') ?? attr(directIndent, 'firstLine'),
  );
  const directHangingIndent = twipToPx(
    attr(directIndent, 'w:hanging') ?? attr(directIndent, 'hanging'),
  );
  const usesNumberingGeometry = Boolean(
    numberPrefix?.suffix === 'tab' &&
      numberPrefix.textStart !== undefined &&
      numberPrefix.hanging !== undefined &&
      numberPrefix.hanging > 0 &&
      directFirstLineIndent === undefined &&
      (directHangingIndent === undefined || directHangingIndent > 0),
  );
  // Word 的优先级是直接段落缩进 > 编号级别缩进 > 段落样式缩进。
  const numberingIndentLeft = usesNumberingGeometry
    ? directIndentLeft ?? numberPrefix?.textStart ?? style.indentLeft
    : style.indentLeft;
  const numberingFirstLineIndent = usesNumberingGeometry
    ? directHangingIndent !== undefined
      ? -directHangingIndent
      : -(numberPrefix?.hanging ?? 0)
    : style.firstLineIndent;
  const numberingTextStyle =
    inlines.find(
      (inline): inline is Extract<DocxInline, { type: 'text' }> =>
        inline.type === 'text' && Boolean(inline.text),
    )?.style ?? style.style;
  const numberingHangingWidth = Math.max(0, -(numberingFirstLineIndent ?? 0));
  const numberingAdvanceWidth = usesNumberingGeometry
    ? numberingHangingWidth
    : undefined;
  const numberingGapWidth = usesNumberingGeometry
    ? resolveDocxNumberingGapWidth(
        numberPrefix?.text ?? '',
        numberingHangingWidth,
        numberingTextStyle?.fontSize ?? 14,
      )
    : 0;
  if (
    numberPrefix &&
    sourceText !== numberPrefix.text &&
    !sourceText.startsWith(`${numberPrefix.text} `) &&
    !sourceText.startsWith(`${numberPrefix.text}\t`)
  ) {
    inlines.unshift({
      type: 'text',
      text: `${numberPrefix.text}${numberPrefix.suffix === 'space' ? ' ' : ''}`,
      // 项目符号通常依赖 numbering.xml 中的 Wingdings/Symbol 字体，不能套用正文中文字体。
      style: numberPrefix.fontFamily
        ? { ...numberingTextStyle, fontFamily: numberPrefix.fontFamily }
        : numberingTextStyle,
      // 编号标记与正文共享同一悬挂缩进，固定前进宽度可让换行正文继续对齐。
      advanceWidth: numberingAdvanceWidth,
    });
    if (numberPrefix.suffix === 'tab') {
      if (usesNumberingGeometry && numberingGapWidth > 0) {
        inlines.splice(1, 0, {
          type: 'text',
          text: '',
          style: numberingTextStyle,
          advanceWidth: numberingGapWidth,
        });
      } else if (!usesNumberingGeometry) {
        inlines.splice(1, 0, { type: 'tab', style: style.style });
      }
    }
  }
  const text = textFromInlines(inlines).trim();
  const outlineLevel =
    text &&
    !options?.insideTable &&
    !options?.insidePageRegion &&
    !style.isTocStyle &&
    style.outlineLevel !== undefined &&
    style.outlineLevel <= 8
      ? style.outlineLevel
      : undefined;
  const visibleRunFontSize = inlines.reduce((maximum, inline) => {
    if (inline.type !== 'text' || !inline.text) return maximum;
    return Math.max(maximum, inline.style?.fontSize ?? 0);
  }, 0);
  const visibleRunLineMetricStyle =
    inlines.find(
      (inline): inline is Extract<DocxInline, { type: 'text' }> =>
        inline.type === 'text' &&
        Boolean(inline.text) &&
        inline.advanceWidth === undefined,
    )?.style ??
    inlines.find(
      (inline): inline is Extract<DocxInline, { type: 'text' }> =>
        inline.type === 'text' && Boolean(inline.text),
    )?.style;
  const visibleRunLineBoxFontSize =
    visibleRunLineMetricStyle?.lineBoxFontSize ??
    visibleRunLineMetricStyle?.fontSize ??
    0;
  // 段落标记格式只影响空段落；存在可见 run 时，文字与行盒均以 run 为准。
  const paragraphMarkFontSize =
    style.paragraphMarkStyle?.fontSize ?? style.style?.fontSize ?? 0;
  const paragraphMarkLineBoxFontSize =
    style.paragraphMarkStyle?.lineBoxFontSize ?? paragraphMarkFontSize;
  const fontSize = visibleRunFontSize || paragraphMarkFontSize || 14;
  const lineBoxFontSize = visibleRunFontSize
    ? visibleRunLineBoxFontSize || fontSize
    : Math.max(fontSize, paragraphMarkLineBoxFontSize);
  const paragraphTextStyle =
    visibleRunFontSize > 0
      ? style.style?.fontSize !== fontSize
        ? { ...style.style, fontSize }
        : style.style
      : style.paragraphMarkStyle ?? style.style;
  // auto 行距必须采用可见 run 或空段落标记的实际字号，否则大字号标题会被正文行盒裁压。
  const paragraphLineHeightFontSize = Math.max(
    DOCX_AUTO_LINE_HEIGHT_BASE_FONT_SIZE_PX,
    visibleRunFontSize ? lineBoxFontSize : paragraphMarkLineBoxFontSize,
    style.style?.fontSize ?? 0,
  );
  const explicitLineHeightPx =
    style.lineHeight === undefined
      ? undefined
      : style.lineHeight <= 4
      ? paragraphLineHeightFontSize * style.lineHeight
      : style.lineHeight;
  // WPS 的纯浮动锚点段落仍按段落标记的字体场景占用网格，不能只使用浏览器默认行盒。
  const usesAnchorYaHeiLineMetrics =
    !visibleRunFontSize &&
    !options?.insideShape &&
    containsOnlyPositionedInlines(inlines) &&
    Boolean(paragraphTextStyle?.bold) &&
    DOCX_YAHEI_FONT_PATTERN.test(paragraphTextStyle?.fontFamily ?? '');
  // 正文与文本框使用不同的默认行盒倍率，但都需按 Word 规则向上吸附文档网格。
  // 网格占用必须按实际可见 run 的字体度量判断；段落标记字体只作为空段落回退。
  const lineMetricFontFamily =
    visibleRunLineMetricStyle?.fontFamily ??
    paragraphTextStyle?.fontFamily ??
    '';
  const defaultLineHeightMultiplier = options?.insideShape
    ? DOCX_SHAPE_LINE_HEIGHT_MULTIPLIER
    : usesAnchorYaHeiLineMetrics
    ? DOCX_ANCHOR_YAHEI_LINE_HEIGHT_MULTIPLIER
    : options?.insideTable &&
      DOCX_COMPACT_SERIF_FONT_PATTERN.test(lineMetricFontFamily)
    ? DOCX_COMPACT_SERIF_LINE_HEIGHT_MULTIPLIER
    : DOCX_BODY_LINE_HEIGHT_MULTIPLIER;
  const naturalLineHeight = lineBoxFontSize * defaultLineHeightMultiplier;
  const resolvedAutoLineHeight =
    style.lineHeightRule === 'auto' &&
    style.lineHeight !== undefined &&
    style.lineHeight <= 4
      ? options?.insideTable
        ? // 表格单倍行距仍采用字体自然行盒，避免单元格文字和自动行高被压扁。
          naturalLineHeight * style.lineHeight
        : explicitLineHeightPx ?? naturalLineHeight
      : Math.max(explicitLineHeightPx ?? 0, naturalLineHeight);
  const gridCandidateLineHeight =
    style.lineHeightRule === 'auto' && explicitLineHeightPx !== undefined
      ? explicitLineHeightPx
      : resolvedAutoLineHeight;
  const snappedDocumentLineHeight =
    style.snapToGrid !== false && context.documentGridLineHeight !== undefined
      ? Math.ceil(gridCandidateLineHeight / context.documentGridLineHeight) *
        context.documentGridLineHeight
      : undefined;
  const autoGridLineHeight =
    style.lineHeightRule === 'auto' &&
    explicitLineHeightPx !== undefined &&
    context.defaultLineHeight !== undefined
      ? visibleRunFontSize
        ? // 可见文字按正文行高的整数倍占格，空段落标记则保留自身声明的自动行距。
          Math.ceil(explicitLineHeightPx / context.defaultLineHeight) *
          context.defaultLineHeight
        : Math.max(explicitLineHeightPx, context.defaultLineHeight)
      : style.lineHeightRule === 'auto' &&
        explicitLineHeightPx !== undefined &&
        context.documentGridLineHeight !== undefined
      ? // 未声明网格类型的 WPS 文档只把正文提升到单格，大字号不应被强制翻倍。
        Math.max(explicitLineHeightPx, context.documentGridLineHeight)
      : undefined;
  const gridLineHeight = options?.insideShape
    ? snappedDocumentLineHeight
    : options?.insideTable
    ? undefined
    : autoGridLineHeight ??
      context.defaultLineHeight ??
      snappedDocumentLineHeight;
  const lineHeight =
    style.lineHeightRule === 'exact' && style.lineHeight !== undefined
      ? style.lineHeight
      : options?.insideTable && style.snapToGrid !== false
      ? resolvedAutoLineHeight
      : style.lineHeightRule === 'auto'
      ? style.snapToGrid !== false &&
        explicitLineHeightPx !== undefined &&
        gridLineHeight !== undefined
        ? gridLineHeight
        : resolvedAutoLineHeight
      : style.snapToGrid !== false &&
        explicitLineHeightPx !== undefined &&
        gridLineHeight !== undefined &&
        explicitLineHeightPx < gridLineHeight
      ? gridLineHeight
      : style.lineHeight === undefined && style.keepNext && style.keepLines
      ? // Word 标题即使未声明行距，也使用独立的字体自然行盒；浏览器 normal 会把标题高度压低。
        lineBoxFontSize * DOCX_HEADING_LINE_HEIGHT_MULTIPLIER
      : style.lineHeight === undefined &&
        style.snapToGrid !== false &&
        gridLineHeight !== undefined
      ? // 未声明行距的普通正文仍需占用默认文档网格，否则连续段落会被压缩到浏览器自然行高。
        gridLineHeight
      : style.lineHeight;

  return {
    id,
    type: 'paragraph',
    inlines,
    text,
    outlineLevel,
    isTableOfContents: style.isTocStyle || undefined,
    keepNext: style.keepNext || undefined,
    keepLines: style.keepLines || undefined,
    // OOXML 未声明时沿用 Word 默认开启的孤行控制。
    widowControl: style.widowControl !== false,
    paragraphStyleId: style.styleId,
    contextualSpacing: style.contextualSpacing || undefined,
    autoSpaceLatin: style.autoSpaceLatin,
    autoSpaceNumber: style.autoSpaceNumber,
    tabStops: style.tabStops,
    align: style.align,
    lineHeight,
    style: paragraphTextStyle,
    spacingBefore: style.spacingBefore,
    spacingAfter: style.spacingAfter,
    indentLeft: numberingIndentLeft,
    indentRight: style.indentRight,
    firstLineIndent: numberingFirstLineIndent,
    backgroundColor: style.backgroundColor,
    borderTop: style.borderTop,
    borderRight: style.borderRight,
    borderBottom: style.borderBottom,
    borderLeft: style.borderLeft,
    paddingTop: style.paddingTop,
    paddingRight: style.paddingRight,
    paddingBottom: style.paddingBottom,
    paddingLeft: style.paddingLeft,
  };
}

function readCellMargins(tcPr: Element | null | undefined) {
  const tcMar =
    childByLocalName(tcPr, 'tcMar') ?? childByLocalName(tcPr, 'tblCellMar');
  const readMargin = (name: string) => {
    const node = childByLocalName(tcMar, name);
    return positiveTwipToPx(attr(node, 'w:w') ?? attr(node, 'w'));
  };
  return {
    paddingTop: readMargin('top'),
    paddingRight: readMargin('right'),
    paddingBottom: readMargin('bottom'),
    paddingLeft: readMargin('left'),
  };
}

/** 合并 `mergeCellMargins` 接收的多份数据。 */
function mergeCellMargins(
  base: Pick<
    DocxTableCell,
    'paddingTop' | 'paddingRight' | 'paddingBottom' | 'paddingLeft'
  >,
  next: Pick<
    DocxTableCell,
    'paddingTop' | 'paddingRight' | 'paddingBottom' | 'paddingLeft'
  >,
) {
  return {
    paddingTop: next.paddingTop ?? base.paddingTop,
    paddingRight: next.paddingRight ?? base.paddingRight,
    paddingBottom: next.paddingBottom ?? base.paddingBottom,
    paddingLeft: next.paddingLeft ?? base.paddingLeft,
  };
}

function readCellBorders(tcPr: Element | null | undefined) {
  const tcBorders = childByLocalName(tcPr, 'tcBorders');
  const top = childByLocalName(tcBorders, 'top');
  const right = childByLocalName(tcBorders, 'right');
  const bottom = childByLocalName(tcBorders, 'bottom');
  const left = childByLocalName(tcBorders, 'left');
  return {
    borderTop: readBorder(top),
    borderRight: readBorder(right),
    borderBottom: readBorder(bottom),
    borderLeft: readBorder(left),
    hasBorderTop: Boolean(top),
    hasBorderRight: Boolean(right),
    hasBorderBottom: Boolean(bottom),
    hasBorderLeft: Boolean(left),
  };
}

function readCellStyle(
  tcNode: Element,
  defaultMargins: Pick<
    DocxTableCell,
    'paddingTop' | 'paddingRight' | 'paddingBottom' | 'paddingLeft'
  >,
  theme: OfficeTheme,
): Omit<DocxTableCell, 'id' | 'blocks'> {
  const tcPr = childByLocalName(tcNode, 'tcPr');
  const gridSpan = childByLocalName(tcPr, 'gridSpan');
  const width = childByLocalName(tcPr, 'tcW');
  const vAlign =
    attr(childByLocalName(tcPr, 'vAlign'), 'w:val') ??
    attr(childByLocalName(tcPr, 'vAlign'), 'val');
  const shading = childByLocalName(tcPr, 'shd');
  const margins = mergeCellMargins(defaultMargins, readCellMargins(tcPr));
  return {
    colSpan: Number(attr(gridSpan, 'w:val') ?? attr(gridSpan, 'val') ?? 1),
    width: twipToPx(attr(width, 'w:w') ?? attr(width, 'w')),
    verticalAlign:
      vAlign === 'center' ? 'middle' : vAlign === 'bottom' ? 'bottom' : 'top',
    backgroundColor: readShading(shading, theme),
    noWrap: readOnOff(childByLocalName(tcPr, 'noWrap')),
    ...readCellBorders(tcPr),
    ...margins,
  };
}

function readCellVerticalMerge(tcNode: Element) {
  const tcPr = childByLocalName(tcNode, 'tcPr');
  const vMerge = childByLocalName(tcPr, 'vMerge');
  if (!vMerge) return undefined;
  const value = readVal(vMerge);
  return value === 'restart' ? 'restart' : 'continue';
}

function readTableRowHeightMultiplier(rowNode: Element) {
  return childrenByLocalName(rowNode, 'tc').reduce(
    (maxMultiplier, cellNode) => {
      const paragraphs = childrenByLocalName(cellNode, 'p');
      const hasPaddingParagraph =
        paragraphs.length > 1 &&
        paragraphs.some((paragraph) => !textContent(paragraph).trim());
      return hasPaddingParagraph
        ? Math.max(maxMultiplier, paragraphs.length)
        : maxMultiplier;
    },
    1,
  );
}

function readTableRowHeight(
  rowNode: Element,
  applyGridHeight: boolean,
): Pick<DocxTableRow, 'cantSplit' | 'height' | 'heightRule'> {
  const trPr = childByLocalName(rowNode, 'trPr');
  const trHeight = childByLocalName(trPr, 'trHeight');
  const height = positiveTwipToPx(
    attr(trHeight, 'w:val') ?? attr(trHeight, 'val'),
  );
  const heightRule = attr(trHeight, 'w:hRule') ?? attr(trHeight, 'hRule');
  const heightMultiplier =
    height !== undefined && height < 80
      ? readTableRowHeightMultiplier(rowNode)
      : 1;
  const lineGridMultiplier =
    applyGridHeight && heightRule === 'atLeast' ? 1.4 : 1;
  return {
    cantSplit: readOnOff(childByLocalName(trPr, 'cantSplit')) || undefined,
    // WPS 的 atLeast 行高仍会叠加正文行网格，直接作为 CSS 最小高度会让表格明显偏扁。
    height:
      height === undefined
        ? undefined
        : height * heightMultiplier * lineGridMultiplier,
    heightRule:
      heightRule === 'exact' || heightRule === 'atLeast'
        ? heightRule
        : height
        ? 'atLeast'
        : undefined,
  };
}

function readCellBlocks(
  cellNode: Element,
  id: string,
  context: ParseContext,
  paragraphContextStyle?: DocxTextStyle,
) {
  const blocks = readDocxBlockChildren(cellNode, id, context, {
    insideTable: true,
    paragraphContextStyle,
  });
  const defaultLineHeight = context.defaultLineHeight;
  const visibleParagraphs = blocks.filter(
    (block): block is DocxParagraphBlock =>
      block.type === 'paragraph' && Boolean(block.text.trim()),
  );
  if (visibleParagraphs.length <= 1 || defaultLineHeight === undefined) {
    return blocks;
  }
  // 正文多段单元格提升到正文网格；已吸附原始文档网格的紧凑代码文本保持较小行距。
  return blocks.map((block) => {
    if (block.type !== 'paragraph' || !block.text.trim()) return block;
    const fontSize = block.style?.fontSize ?? 14;
    const explicitLineHeight =
      block.lineHeight === undefined
        ? undefined
        : block.lineHeight > 4
        ? block.lineHeight
        : fontSize * block.lineHeight;
    const followsDocumentGrid =
      block.lineHeight !== undefined &&
      block.lineHeight > 4 &&
      context.documentGridLineHeight !== undefined &&
      Math.abs(block.lineHeight - context.documentGridLineHeight) < 0.5;
    const followsBodyGrid =
      explicitLineHeight === undefined || !followsDocumentGrid;
    return followsBodyGrid &&
      (explicitLineHeight === undefined ||
        explicitLineHeight < defaultLineHeight)
      ? {
          ...block,
          // 表格 flex 行盒会把绝对行高向上取整约 1px，解析阶段抵消以保持文档网格总高。
          lineHeight: Math.max(1, defaultLineHeight - 1),
        }
      : block;
  });
}

/** 判断单元格文本是否已由绝对文档网格行高完整约束。 */
function usesDocumentGridCellPadding(blocks: DocxBlock[]) {
  const paragraphs = blocks.filter(
    (block): block is DocxParagraphBlock =>
      block.type === 'paragraph' && Boolean(block.text.trim()),
  );
  return (
    paragraphs.length > 0 &&
    paragraphs.every(
      (paragraph) =>
        paragraph.lineHeight !== undefined && paragraph.lineHeight > 4,
    )
  );
}

/** 获取 `getParagraphAnchorLineHeight` 返回的数据。 */
function getParagraphAnchorLineHeight(block: DocxParagraphBlock) {
  const fontSize = block.style?.fontSize ?? 14;
  if (block.lineHeight === undefined) return fontSize * 1.2;
  return block.lineHeight > 4 ? block.lineHeight : fontSize * block.lineHeight;
}

function isPositionedOnlyParagraph(
  block: DocxBlock | undefined,
): block is DocxParagraphBlock {
  return Boolean(
    block?.type === 'paragraph' && containsOnlyPositionedInlines(block.inlines),
  );
}

function offsetTableAfterPositionedParagraph(
  table: DocxTableBlock,
  previousBlock: DocxBlock | undefined,
) {
  if (!isPositionedOnlyParagraph(previousBlock)) return table;
  const lineHeight = getParagraphAnchorLineHeight(previousBlock);
  if (!table.position) {
    return {
      ...table,
      // 锚点段落自身已占用一个正文行盒；流式表格应紧接该行，不能再叠加通用表格留白。
      marginTop: 0,
    };
  }
  if (table.position.relativeFromV !== 'text') return table;
  return {
    ...table,
    position: {
      ...table.position,
      top: table.position.top + lineHeight,
    },
  };
}

function readTableWidth(tblW: Element | null | undefined, columns: number[]) {
  const widthType = attr(tblW, 'w:type') ?? attr(tblW, 'type');
  if (widthType === 'pct' && columns.length) {
    return (
      columns.reduce((sum, width) => sum + width, 0) +
      DOCX_DEFAULT_TABLE_EDGE_OFFSET * 2
    );
  }
  return positiveTwipToPx(attr(tblW, 'w:w') ?? attr(tblW, 'w'));
}

/** 将输入标准化为 `normalizeTableForBlockContext` 返回的结构。 */
function normalizeTableForBlockContext(
  table: DocxTableBlock,
  options?: ReadBlockChildrenOptions,
) {
  if (!options?.insideShape || !table.position) return table;
  return {
    ...table,
    // 文本框已经承载了页面锚点，内部表格再使用 tblpPr 会把页面坐标叠加一次。
    position: undefined,
    insideShape: true,
    visualOffsetTop: undefined,
  };
}

function readTablePosition(
  tblPr: Element | null | undefined,
): DocxPosition | undefined {
  const tblpPr = childByLocalName(tblPr, 'tblpPr');
  if (!tblpPr) return undefined;

  const rawLeft = twipToPx(attr(tblpPr, 'w:tblpX') ?? attr(tblpPr, 'tblpX'));
  const rawTop = twipToPx(attr(tblpPr, 'w:tblpY') ?? attr(tblpPr, 'tblpY'));
  if (rawLeft === undefined || rawTop === undefined) return undefined;

  const leftFromText =
    twipToPx(attr(tblpPr, 'w:leftFromText') ?? attr(tblpPr, 'leftFromText')) ??
    0;
  const horzAnchor = attr(tblpPr, 'w:horzAnchor') ?? attr(tblpPr, 'horzAnchor');
  const vertAnchor = attr(tblpPr, 'w:vertAnchor') ?? attr(tblpPr, 'vertAnchor');

  return {
    left: Math.round(rawLeft - leftFromText),
    top: Math.round(rawTop),
    relativeFromH:
      horzAnchor === 'page'
        ? 'margin'
        : (horzAnchor as DocxPosition['relativeFromH']),
    relativeFromV:
      vertAnchor === 'text'
        ? 'text'
        : (vertAnchor as DocxPosition['relativeFromV']),
  };
}

function parseTable(
  tblNode: Element,
  id: string,
  context: ParseContext,
): DocxTableBlock {
  const tblPr = childByLocalName(tblNode, 'tblPr');
  const tblW = childByLocalName(tblPr, 'tblW');
  const tableStyleId =
    readVal(childByLocalName(tblPr, 'tblStyle')) ??
    context.styles.defaults.tableStyleId;
  const tableParagraphStyle = resolveDocxStyle(tableStyleId, context.styles);
  const align = mapAlignment(readVal(childByLocalName(tblPr, 'jc')));
  const columns = childrenByLocalName(
    childByLocalName(tblNode, 'tblGrid'),
    'gridCol',
  )
    .map((col) => positiveTwipToPx(attr(col, 'w:w') ?? attr(col, 'w')))
    .filter((width): width is number => width !== undefined);
  const tableMargins = mergeCellMargins(
    resolveDocxTableCellMargins(tableStyleId, context.styles) ?? {},
    readCellMargins(tblPr),
  );
  const styleBorders = resolveDocxTableBorders(tableStyleId, context.styles);
  const directBorders = readDocxTableBorders(tblPr);
  const tableBorders = {
    top: directBorders.top ?? styleBorders?.top,
    right: directBorders.right ?? styleBorders?.right,
    bottom: directBorders.bottom ?? styleBorders?.bottom,
    left: directBorders.left ?? styleBorders?.left,
    insideHorizontal:
      directBorders.insideHorizontal ?? styleBorders?.insideHorizontal,
    insideVertical:
      directBorders.insideVertical ?? styleBorders?.insideVertical,
  };
  const result: DocxTableBlock = {
    id,
    type: 'table',
    width: readTableWidth(tblW, columns),
    align: align === 'center' || align === 'right' ? align : 'left',
    columns,
    position: readTablePosition(tblPr),
    rows: [],
  };

  const activeVerticalMerges = new Map<
    number,
    {
      /** 当前纵向合并链保存的表格单元格。 */
      cell: DocxTableCell;
      /** 表格单元格横向跨越的列数。 */
      colSpan: number;
    }
  >();
  result.rows = childrenByLocalName(tblNode, 'tr').map((rowNode, rowIndex) => {
    let columnIndex = 0;
    const cells: DocxTableCell[] = [];

    childrenByLocalName(rowNode, 'tc').forEach((cellNode, cellIndex) => {
      const verticalMerge = readCellVerticalMerge(cellNode);
      const cellId = `${id}-cell-${rowIndex + 1}-${cellIndex + 1}`;
      const cellStyle = readCellStyle(cellNode, tableMargins, context.theme);
      const colSpan =
        cellStyle.colSpan && cellStyle.colSpan > 1 ? cellStyle.colSpan : 1;

      if (verticalMerge === 'continue') {
        const activeMerge = activeVerticalMerges.get(columnIndex);
        if (activeMerge) {
          activeMerge.cell.rowSpan = (activeMerge.cell.rowSpan ?? 1) + 1;
          columnIndex += activeMerge.colSpan;
          return;
        }
      } else {
        activeVerticalMerges.delete(columnIndex);
      }

      const blocks = readCellBlocks(
        cellNode,
        cellId,
        context,
        tableParagraphStyle,
      );
      const gridControlsVerticalSpacing = usesDocumentGridCellPadding(blocks);
      const cell: DocxTableCell = {
        id: cellId,
        ...cellStyle,
        // 文档网格已提供完整行盒时不再叠加浏览器补偿内边距。
        paddingTop:
          gridControlsVerticalSpacing && cellStyle.paddingTop === undefined
            ? 0
            : cellStyle.paddingTop,
        paddingBottom:
          gridControlsVerticalSpacing && cellStyle.paddingBottom === undefined
            ? 0
            : cellStyle.paddingBottom,
        blocks,
      };
      cells.push(cell);

      if (verticalMerge === 'restart') {
        cell.rowSpan = 1;
        activeVerticalMerges.set(columnIndex, { cell, colSpan });
      }

      columnIndex += colSpan;
    });

    return {
      id: `${id}-row-${rowIndex + 1}`,
      ...readTableRowHeight(rowNode, context.defaultLineHeight !== undefined),
      cells,
    };
  });
  result.rows.forEach((row, rowIndex) => {
    row.cells.forEach((cell, cellIndex) => {
      const isFirstRow = rowIndex === 0;
      const isLastRow = rowIndex === result.rows.length - 1;
      const isFirstCell = cellIndex === 0;
      const isLastCell = cellIndex === row.cells.length - 1;
      if (!cell.hasBorderTop && !cell.borderTop) {
        cell.borderTop = isFirstRow
          ? tableBorders.top
          : tableBorders.insideHorizontal;
      }
      if (!cell.hasBorderBottom && !cell.borderBottom) {
        cell.borderBottom = isLastRow
          ? tableBorders.bottom
          : tableBorders.insideHorizontal;
      }
      if (!cell.hasBorderLeft && !cell.borderLeft) {
        cell.borderLeft = isFirstCell
          ? tableBorders.left
          : tableBorders.insideVertical;
      }
      if (!cell.hasBorderRight && !cell.borderRight) {
        cell.borderRight = isLastCell
          ? tableBorders.right
          : tableBorders.insideVertical;
      }
    });
  });

  return result;
}

/** 递归读取 DOCX 容器节点中的段落和表格。 */
export function readDocxBlockChildren(
  node: Element | null | undefined,
  id: string,
  context: ParseContext,
  options?: ReadBlockChildrenOptions,
): DocxBlock[] {
  const blocks: DocxBlock[] = [];
  let paragraphIndex = 0;
  let tableIndex = 0;

  Array.from(node?.children ?? []).forEach((child) => {
    if (matchesLocalName(child, 'p')) {
      paragraphIndex += 1;
      blocks.push(
        ...readParagraphBlocks(
          child,
          `${id}-p-${paragraphIndex}`,
          context,
          options,
        ),
      );
    }
    if (matchesLocalName(child, 'tbl')) {
      tableIndex += 1;
      const table = normalizeTableForBlockContext(
        parseTable(child, `${id}-table-${tableIndex}`, context),
        options,
      );
      blocks.push(
        offsetTableAfterPositionedParagraph(table, blocks[blocks.length - 1]),
      );
    }
  });

  return blocks;
}

function readSectionPage(sectPr: Element | null | undefined): DocxPage {
  const pgSz = childByLocalName(sectPr, 'pgSz');
  const pgMar = childByLocalName(sectPr, 'pgMar');
  const pgBorders = childByLocalName(sectPr, 'pgBorders');

  return {
    width: Math.round(
      twipToPx(attr(pgSz, 'w:w') ?? attr(pgSz, 'w')) ?? DEFAULT_DOCX_PAGE.width,
    ),
    minHeight: Math.round(
      twipToPx(attr(pgSz, 'w:h') ?? attr(pgSz, 'h')) ??
        DEFAULT_DOCX_PAGE.minHeight,
    ),
    marginTop: Math.round(
      twipToPx(attr(pgMar, 'w:top') ?? attr(pgMar, 'top')) ??
        DEFAULT_DOCX_PAGE.marginTop,
    ),
    marginRight: Math.round(
      twipToPx(attr(pgMar, 'w:right') ?? attr(pgMar, 'right')) ??
        DEFAULT_DOCX_PAGE.marginRight,
    ),
    marginBottom: Math.round(
      twipToPx(attr(pgMar, 'w:bottom') ?? attr(pgMar, 'bottom')) ??
        DEFAULT_DOCX_PAGE.marginBottom,
    ),
    marginLeft: Math.round(
      twipToPx(attr(pgMar, 'w:left') ?? attr(pgMar, 'left')) ??
        DEFAULT_DOCX_PAGE.marginLeft,
    ),
    headerDistance: twipToPx(attr(pgMar, 'w:header') ?? attr(pgMar, 'header')),
    footerDistance: twipToPx(attr(pgMar, 'w:footer') ?? attr(pgMar, 'footer')),
    borderTop: readBorder(childByLocalName(pgBorders, 'top')),
    borderRight: readBorder(childByLocalName(pgBorders, 'right')),
    borderBottom: readBorder(childByLocalName(pgBorders, 'bottom')),
    borderLeft: readBorder(childByLocalName(pgBorders, 'left')),
  };
}

function readPage(bodyNode: Element | null | undefined): DocxPage {
  return readSectionPage(childByLocalName(bodyNode, 'sectPr'));
}

/** 获取 OOXML 部件对应的关系文件路径。 */
function getPartRelationshipsPath(partPath: string) {
  const lastSlash = partPath.lastIndexOf('/');
  const directory = lastSlash >= 0 ? partPath.slice(0, lastSlash) : '';
  const fileName = lastSlash >= 0 ? partPath.slice(lastSlash + 1) : partPath;
  return `${directory ? `${directory}/` : ''}_rels/${fileName}.rels`;
}

/** 解析页眉部件，并同步其中新增的媒体和对象索引。 */
function readHeaderPartBlocks(
  partPath: string,
  type: string,
  context: ParseContext,
) {
  const xml = readXml(context.packageState.entries, partPath);
  if (!xml) return undefined;
  const partContext: ParseContext = {
    ...context,
    documentRels:
      context.packageState.relationships[getPartRelationshipsPath(partPath)] ??
      {},
  };
  const root = parseXml(xml).documentElement;
  const blocks = readDocxBlockChildren(root, `header-${type}`, partContext, {
    insidePageRegion: true,
  });
  context.imageIndex = partContext.imageIndex;
  context.chartIndex = partContext.chartIndex;
  context.shapeIndex = partContext.shapeIndex;
  return blocks;
}

/** 读取当前节的页眉、页脚页码及首页差异设置。 */
export function readDocxSectionPageRegions(
  sectPr: Element | null | undefined,
  context: ParseContext,
): Pick<
  DocxPageContent,
  'headers' | 'footerPageNumbers' | 'differentFirstPage'
> {
  const headers: DocxPageRegionVariants<DocxBlock[]> = {};
  const footerPageNumbers: DocxPageRegionVariants<boolean> = {};
  childrenByLocalName(sectPr, 'headerReference').forEach((reference) => {
    const type = (attr(reference, 'w:type') ?? attr(reference, 'type')) as
      | 'default'
      | 'first'
      | 'even';
    const relationshipId = attr(reference, 'r:id') ?? attr(reference, 'id');
    const partPath = relationshipId
      ? context.documentRels[relationshipId]?.target
      : undefined;
    if (!partPath || !type) return;
    const blocks = readHeaderPartBlocks(partPath, type, context);
    if (blocks?.length) headers[type] = blocks;
  });
  childrenByLocalName(sectPr, 'footerReference').forEach((reference) => {
    const type = (attr(reference, 'w:type') ?? attr(reference, 'type')) as
      | 'default'
      | 'first'
      | 'even';
    const relationshipId = attr(reference, 'r:id') ?? attr(reference, 'id');
    const partPath = relationshipId
      ? context.documentRels[relationshipId]?.target
      : undefined;
    const xml = partPath
      ? readXml(context.packageState.entries, partPath)
      : undefined;
    if (
      type &&
      xml &&
      /\bPAGE\b/i.test(textContent(parseXml(xml).documentElement))
    ) {
      footerPageNumbers[type] = true;
    }
  });
  return {
    headers: Object.keys(headers).length ? headers : undefined,
    footerPageNumbers: Object.keys(footerPageNumbers).length
      ? footerPageNumbers
      : undefined,
    differentFirstPage: Boolean(childByLocalName(sectPr, 'titlePg')),
  };
}

function markTitle(blocks: DocxBlock[]) {
  const firstParagraph = blocks.find(
    (block): block is DocxParagraphBlock =>
      block.type === 'paragraph' && Boolean(block.text),
  );
  return firstParagraph?.text ?? 'DOCX 文档';
}

/** 为缺少显式段前距的首个大字号居中标题恢复 Word 封面的视觉留白。 */
export function applyDocxCoverTitleSpacing(blocks: DocxBlock[]) {
  const firstParagraph = blocks.find(
    (block): block is DocxParagraphBlock =>
      block.type === 'paragraph' && Boolean(block.text),
  );
  const fontSize = firstParagraph?.style?.fontSize ?? 0;
  if (
    firstParagraph &&
    firstParagraph.align === 'center' &&
    fontSize >= 28 &&
    firstParagraph.spacingBefore === undefined
  ) {
    // Word 的大字号空段行框高于浏览器默认行框，用字号比例补足封面标题前的差值。
    firstParagraph.spacingBefore = Math.round(fontSize * 0.5);
  }
}

function isEmptySpacerParagraph(block: DocxBlock) {
  return (
    block.type === 'paragraph' &&
    !block.text &&
    !block.inlines.length &&
    !block.backgroundColor
  );
}

function hasRenderableBlockContent(block: DocxBlock) {
  if (block.type === 'paragraph')
    return Boolean(block.text || block.inlines.length);
  if (block.type === 'table')
    return block.rows.some((row) =>
      row.cells.some((cell) => cell.blocks.length),
    );
  return true;
}

function isFullPagePositionedShape(
  position: DocxPosition | undefined,
  size: {
    /** 对象宽度，单位为标准化渲染像素。 */
    width?: number;
    /** 对象高度，单位为标准化渲染像素。 */
    height?: number;
  },
  page: DocxPage,
) {
  if (!position || !size.width || !size.height) return false;
  return (
    size.width >= page.width * 0.85 && size.height >= page.minHeight * 0.75
  );
}

function blockHasFullPagePositionedShape(block: DocxBlock, page: DocxPage) {
  if (block.type === 'chart') {
    return isFullPagePositionedShape(block.position, block, page);
  }

  if (block.type !== 'paragraph') return false;

  return block.inlines.some((inline) => {
    if (inline.type === 'image')
      return isFullPagePositionedShape(
        inline.image.position,
        inline.image,
        page,
      );
    if (inline.type === 'shape')
      return isFullPagePositionedShape(
        inline.shape.position,
        inline.shape,
        page,
      );
    if (inline.type === 'chart')
      return isFullPagePositionedShape(
        inline.chart.position,
        inline.chart,
        page,
      );
    return false;
  });
}

/** 按 `splitSectionOverflowPage` 的规则拆分输入数据。 */
function splitSectionOverflowPage(
  pageContent: DocxPageContent,
): DocxPageContent[] {
  const splitPages: DocxPageContent[] = [];
  let currentBlocks: DocxBlock[] = [];
  let pendingSpacers: DocxBlock[] = [];
  let currentHasContent = false;
  let currentHasFullPageShape = false;
  let didSplit = false;

  const pushCurrentPage = () => {
    if (!currentBlocks.length) return;
    splitPages.push({
      ...pageContent,
      id: `${pageContent.id}-split-${splitPages.length + 1}`,
      blocks: currentBlocks,
    });
    currentBlocks = [];
    pendingSpacers = [];
    currentHasContent = false;
    currentHasFullPageShape = false;
  };

  pageContent.blocks.forEach((block) => {
    if (isEmptySpacerParagraph(block)) {
      pendingSpacers.push(block);
      return;
    }

    const startsWithFullPageShape = blockHasFullPagePositionedShape(
      block,
      pageContent.page,
    );
    if (
      startsWithFullPageShape &&
      currentHasFullPageShape &&
      currentHasContent &&
      pendingSpacers.length >= 2
    ) {
      // WPS 会把连续页面放在同一个 section 中，第二个整页背景通常就是新的自然分页。
      pushCurrentPage();
      didSplit = true;
    } else if (pendingSpacers.length) {
      currentBlocks.push(...pendingSpacers);
      pendingSpacers = [];
    }

    currentBlocks.push(block);
    currentHasFullPageShape =
      currentHasFullPageShape || startsWithFullPageShape;
    currentHasContent = currentHasContent || hasRenderableBlockContent(block);
  });

  if (!didSplit) return [pageContent];

  pushCurrentPage();
  return splitPages.length ? splitPages : [pageContent];
}

/** 将输入标准化为 `normalizeDocxPages` 返回的结构。 */
function normalizeDocxPages(pages: DocxPageContent[]) {
  return pages
    .flatMap((pageContent) => splitSectionOverflowPage(pageContent))
    .map((pageContent, index) => ({
      ...pageContent,
      id: `docx-page-${index + 1}`,
      // WPS 相册等模板把整页内容保存为定位画布，这类物理页不应被浏览器测量结果再次拆开。
      preservePhysicalPage:
        pageContent.blocks.some((block) =>
          blockHasFullPagePositionedShape(block, pageContent.page),
        ) || undefined,
    }));
}

/** 判断节点是否包含显式分页符。 */
function hasExplicitPageBreak(node: Element) {
  return descendantsByLocalName(node, 'br').some(
    (breakNode) =>
      (attr(breakNode, 'w:type') ?? attr(breakNode, 'type')) === 'page',
  );
}

/** 判断段落是否只负责承载显式分页符，不应生成可见空行。 */
function isPageBreakOnlyParagraph(node: Element) {
  if (!matchesLocalName(node, 'p') || !hasExplicitPageBreak(node)) return false;
  return (
    !textContent(node).trim() &&
    descendantsByLocalName(node, 'drawing').length === 0 &&
    descendantsByLocalName(node, 'pict').length === 0
  );
}

/** 读取表格内分页标记所在的行索引，分页从该行之前开始。 */
function readTablePageBreakRows(tableNode: Element) {
  return childrenByLocalName(tableNode, 'tr')
    .map((rowNode, rowIndex) => (hasExplicitPageBreak(rowNode) ? rowIndex : -1))
    .filter((rowIndex) => rowIndex >= 0);
}

/** 为物化与流式路径提供同一组 DOCX 块解析规则。 */
export const docxBlockParseOperations: DocxBlockParseOperations<DocxParseContext> =
  {
    // lastRenderedPageBreak 只是上次排版缓存，字体或页面环境改变后可能失效，不能作为硬分页边界。
    hasExplicitPageBreak,
    readTablePageBreakRows,
    isPageBreakOnlyParagraph,
    isParagraph: (node) => matchesLocalName(node, 'p'),
    isTable: (node) => matchesLocalName(node, 'tbl'),
    readParagraphBlocks,
    // 流式大文件解析绕过 readBlockChildren，这里保持与完整物化路径一致。
    parseTable,
    offsetTable: offsetTableAfterPositionedParagraph,
    readParagraphSection: (node) =>
      matchesLocalName(node, 'p')
        ? childByLocalName(childByLocalName(node, 'pPr'), 'sectPr')
        : null,
    readSectionPage,
    readSectionRegions: readDocxSectionPageRegions,
  };

/** 读取合成或完整 body 的默认页面属性。 */
export const readDocxBodyPage = readPage;

/** 标准化物理页 ID，并保留 WPS 整页形状的溢出拆分规则。 */
export const normalizeDocxPageContents = normalizeDocxPages;

/** 从当前已解析块推导文档标题。 */
export const markDocxTitle = markTitle;
