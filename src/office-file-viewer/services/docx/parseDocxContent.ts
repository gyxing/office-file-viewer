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
  inlineInheritedStyle,
  mapAlignment,
  mergeTextStyle,
  positiveTwipToPx,
  readBorder,
  readOnOff,
  readShading,
  readVal,
  resolveParagraphStyle,
  resolveRunStyle,
  twipToPx,
  type DocxParseContext,
  type ReadBlockChildrenOptions,
} from './docxParsingContext';
import type { DocxBlockParseOperations } from './parseDocxBlock';
import { createDocxDrawingParser } from './parseDocxDrawing';
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

// Word 的正文流表格会在文档网格边界后保留约 6pt，浏览器表格需显式补回。
/** DOCX 流式表格相对正文顶部的视觉修正量。 */
const DOCX_FLOW_TABLE_TOP_OFFSET = 8;
// 百分比表宽不包含首尾默认单元格边距，Word 会把两侧各约 7px 绘制到正文边界外。
/** DOCX 表格缺少边缘设置时使用的默认修正量。 */
const DOCX_DEFAULT_TABLE_EDGE_OFFSET = 7;

/** 将单个 run 的文本、换行和绘图子节点转换为行内模型。 */
function parseRun(
  runNode: Element,
  paragraphStyle: DocxTextStyle | undefined,
  context: ParseContext,
): DocxInline[] {
  const runStyle = mergeTextStyle(
    inlineInheritedStyle(paragraphStyle),
    resolveRunStyle(
      childByLocalName(runNode, 'rPr'),
      context.styles,
      context.theme,
    ),
  );
  const inlines: DocxInline[] = [];

  Array.from(runNode.children).forEach((child) => {
    if (matchesLocalName(child, 't')) {
      inlines.push({ type: 'text', text: textContent(child), style: runStyle });
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
    if (drawingInline) inlines.push(drawingInline);
  });

  return inlines;
}

function readParagraphBlocks(
  pNode: Element,
  id: string,
  context: ParseContext,
  options?: ReadBlockChildrenOptions,
): DocxParagraphBlock[] {
  const paragraph = parseParagraph(pNode, id, context, options);
  return [paragraph];
}

/** 按“最终状态”读取段落行内内容：保留插入/移入内容，忽略删除/移出内容。 */
function readParagraphRunChildren(
  parentNode: Element,
  paragraphStyle: DocxTextStyle | undefined,
  context: ParseContext,
) {
  const inlines: DocxInline[] = [];

  Array.from(parentNode.children).forEach((child) => {
    if (matchesLocalName(child, 'r')) {
      inlines.push(...parseRun(child, paragraphStyle, context));
      return;
    }
    if (matchesLocalName(child, 'del') || matchesLocalName(child, 'moveFrom')) {
      return;
    }
    if (matchesLocalName(child, 'sdt')) {
      const content = childByLocalName(child, 'sdtContent');
      if (content) {
        inlines.push(
          ...readParagraphRunChildren(content, paragraphStyle, context),
        );
      }
      return;
    }
    if (
      matchesLocalName(child, 'hyperlink') ||
      matchesLocalName(child, 'ins') ||
      matchesLocalName(child, 'moveTo') ||
      matchesLocalName(child, 'smartTag') ||
      matchesLocalName(child, 'customXml') ||
      matchesLocalName(child, 'fldSimple') ||
      matchesLocalName(child, 'sdtContent')
    ) {
      inlines.push(...readParagraphRunChildren(child, paragraphStyle, context));
    }
  });

  return inlines;
}

function readParagraphRuns(
  pNode: Element,
  paragraphStyle: DocxTextStyle | undefined,
  context: ParseContext,
) {
  return readParagraphRunChildren(pNode, paragraphStyle, context);
}

function textFromInlines(inlines: DocxInline[]) {
  return inlines
    .map((inline) =>
      inline.type === 'text' ? inline.text : inline.type === 'tab' ? '\t' : '',
    )
    .join('');
}

function parseParagraph(
  pNode: Element,
  id: string,
  context: ParseContext,
  options?: ReadBlockChildrenOptions,
): DocxParagraphBlock {
  const pPr = childByLocalName(pNode, 'pPr');
  const style = resolveParagraphStyle(pPr, context.styles, context.theme);
  const inlines = readParagraphRuns(pNode, style.style, context);
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
        ? { ...style.style, fontFamily: numberPrefix.fontFamily }
        : style.style,
    });
    if (numberPrefix.suffix === 'tab')
      inlines.splice(1, 0, { type: 'tab', style: style.style });
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
  const fontSize = style.style?.fontSize ?? 14;
  const explicitLineHeightPx =
    style.lineHeight === undefined
      ? undefined
      : style.lineHeight <= 4
      ? fontSize * style.lineHeight
      : style.lineHeight;
  // Word 会把行盒向上吸附到完整网格；按段落实际字号计算，避免正文被浏览器默认行高压缩。
  const snappedDocumentLineHeight =
    style.snapToGrid !== false && context.documentGridLineHeight !== undefined
      ? Math.ceil(
          (explicitLineHeightPx ?? fontSize * (4 / 3)) /
            context.documentGridLineHeight,
        ) * context.documentGridLineHeight
      : undefined;
  const gridLineHeight = options?.insideTable
    ? context.documentGridLineHeight
    : context.defaultLineHeight ?? snappedDocumentLineHeight;
  const lineHeight =
    style.snapToGrid !== false &&
    gridLineHeight !== undefined &&
    (explicitLineHeightPx === undefined ||
      explicitLineHeightPx < gridLineHeight)
      ? gridLineHeight
      : style.lineHeight;

  return {
    id,
    type: 'paragraph',
    inlines,
    text,
    outlineLevel,
    isTableOfContents: style.isTocStyle || undefined,
    tabStops: style.tabStops,
    align: style.align,
    lineHeight,
    style: style.style,
    spacingBefore: style.spacingBefore,
    spacingAfter: style.spacingAfter,
    indentLeft: style.indentLeft,
    indentRight: style.indentRight,
    firstLineIndent: style.firstLineIndent,
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

/** 读取表格级外框和内部网格线，供未单独声明边框的单元格继承。 */
function readTableBorders(tblPr: Element | null | undefined) {
  const borders = childByLocalName(tblPr, 'tblBorders');
  return {
    top: readBorder(childByLocalName(borders, 'top')),
    right: readBorder(childByLocalName(borders, 'right')),
    bottom: readBorder(childByLocalName(borders, 'bottom')),
    left: readBorder(childByLocalName(borders, 'left')),
    insideHorizontal: readBorder(childByLocalName(borders, 'insideH')),
    insideVertical: readBorder(childByLocalName(borders, 'insideV')),
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
): Pick<DocxTableRow, 'height' | 'heightRule'> {
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

function readCellBlocks(cellNode: Element, id: string, context: ParseContext) {
  const blocks = readDocxBlockChildren(cellNode, id, context, {
    insideTable: true,
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
  if (!block || block.type !== 'paragraph' || !block.inlines.length)
    return false;
  return block.inlines.every((inline) => {
    if (inline.type === 'text') return !inline.text.trim();
    if (inline.type === 'break') return false;
    if (inline.type === 'image') return Boolean(inline.image.position);
    if (inline.type === 'shape') return Boolean(inline.shape.position);
    if (inline.type === 'chart') return Boolean(inline.chart.position);
    return false;
  });
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
      marginTop: (table.marginTop ?? 0) + lineHeight,
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

/** 为顶层正文流表格补回 Word 文档网格在表格边界前保留的留白。 */
function offsetTopLevelFlowTable(table: DocxTableBlock) {
  if (table.position) return table;
  return {
    ...table,
    marginTop: (table.marginTop ?? 0) + DOCX_FLOW_TABLE_TOP_OFFSET,
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
  const align = mapAlignment(readVal(childByLocalName(tblPr, 'jc')));
  const columns = childrenByLocalName(
    childByLocalName(tblNode, 'tblGrid'),
    'gridCol',
  )
    .map((col) => positiveTwipToPx(attr(col, 'w:w') ?? attr(col, 'w')))
    .filter((width): width is number => width !== undefined);
  const tableMargins = readCellMargins(tblPr);
  const tableBorders = readTableBorders(tblPr);
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

      const blocks = readCellBlocks(cellNode, cellId, context);
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
      const flowTable =
        !options?.insideShape &&
        !options?.insideTable &&
        !options?.insidePageRegion
          ? offsetTopLevelFlowTable(table)
          : table;
      blocks.push(
        offsetTableAfterPositionedParagraph(
          flowTable,
          blocks[blocks.length - 1],
        ),
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
    firstParagraph.spacingBefore = Math.round(fontSize * 0.8);
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
    }));
}

/** 判断节点是否包含 Word/WPS 保存的上次渲染分页位置。 */
function hasRenderedPageBreak(node: Element) {
  return descendantsByLocalName(node, 'lastRenderedPageBreak').length > 0;
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
    .map((rowNode, rowIndex) =>
      hasRenderedPageBreak(rowNode) || hasExplicitPageBreak(rowNode)
        ? rowIndex
        : -1,
    )
    .filter((rowIndex) => rowIndex >= 0);
}

/** 为物化与流式路径提供同一组 DOCX 块解析规则。 */
export const docxBlockParseOperations: DocxBlockParseOperations<DocxParseContext> =
  {
    hasRenderedPageBreak,
    hasExplicitPageBreak,
    readTablePageBreakRows,
    isPageBreakOnlyParagraph,
    isParagraph: (node) => matchesLocalName(node, 'p'),
    isTable: (node) => matchesLocalName(node, 'tbl'),
    readParagraphBlocks,
    // 流式大文件解析绕过 readBlockChildren，这里保持与完整物化路径一致。
    parseTable: (node, id, context) =>
      offsetTopLevelFlowTable(parseTable(node, id, context)),
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
