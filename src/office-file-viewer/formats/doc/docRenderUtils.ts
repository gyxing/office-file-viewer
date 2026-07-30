// docRenderUtils 提供 DOC 降级渲染所需的样式转换和图片排版辅助方法。
import type { CSSProperties } from 'react';
import type {
  DocBlock,
  DocImage,
  DocPage,
  DocParagraphBlock,
  DocTextInline,
  DocTextStyle,
} from '../../services/doc/types';

export const DOC_IMAGE_ROW_GAP = 6;

// 页面可用高度已扣除上下页边距；不再重复预留缓冲，避免把本应同页的图片组拆页。
const DOC_PAGE_HEIGHT_BUFFER = 0;
// 浏览器行盒会累计小数像素，允许极小误差可避免 1px 级伪溢出触发整页后移。
const DOC_PAGE_HEIGHT_TOLERANCE = 2;
const DOC_IMAGE_LAYOUT_ROW_GAP = 12;
const DOC_IMAGE_LAYOUT_VERTICAL_MARGIN = 10;

// DOC 没有 OOXML 的显式页面模型，这里记录前端估算分页后的页面块集合。
/** 描述 PaginatedDocPage 在 DOC 渲染中的数据结构。 */
export type PaginatedDocPage = {
  /** PaginatedDocPage 在所属文档或任务中的唯一标识。 */
  id: string;
  /** PaginatedDocPage 包含的 blocks 有序集合。 */
  blocks: DocBlock[];
};

// DOC 解析出的文本样式字段和 React CSS 字段基本一一对应，集中转换便于后续补充新属性。
/** 执行 `docTextStyleToCss` 封装的DOC 渲染处理步骤。 */
export function docTextStyleToCss(style?: DocTextStyle): CSSProperties {
  if (!style) return {};

  const css: CSSProperties = {
    color: style.color,
    background: style.backgroundColor,
    borderColor: style.borderColor,
    borderWidth: style.borderWidth,
    borderStyle: style.borderStyle,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight,
    fontStyle: style.fontStyle,
    textDecoration: style.textDecoration,
    textAlign: style.textAlign,
    // React 的数字 lineHeight 会按无单位倍数处理；DOC 解析出的绝对行距需显式补 px。
    lineHeight:
      style.lineHeight !== undefined && style.lineHeight > 4
        ? `${style.lineHeight}px`
        : style.lineHeight,
    fontFamily: style.fontFamily,
    marginLeft: style.indentLeft,
    marginRight: style.indentRight,
    textIndent: style.firstLineIndent,
    marginTop: style.spacingBefore,
    marginBottom: style.spacingAfter,
    paddingTop: style.paddingTop,
    paddingRight: style.paddingRight,
    paddingBottom: style.paddingBottom,
    paddingLeft: style.paddingLeft,
  };
  return Object.fromEntries(
    Object.entries(css).filter(([, value]) => value !== undefined),
  ) as CSSProperties;
}

/** 执行 `inlineStyleToCss` 封装的DOC 渲染处理步骤。 */
export function inlineStyleToCss(
  style?: DocTextStyle,
  options?: {
    /** 是否保留块级模型自身的字体与段落样式；未提供时使用来源格式或渲染器的默认行为。 */
    preserveBlockTypography?: boolean;
  },
): CSSProperties {
  const css = docTextStyleToCss(style);
  // 行内片段不能继承段落级缩进/间距，否则会把整段排版撑乱。
  delete css.textAlign;
  delete css.marginLeft;
  delete css.marginRight;
  delete css.textIndent;
  delete css.marginTop;
  delete css.marginBottom;
  delete css.paddingTop;
  delete css.paddingRight;
  delete css.paddingBottom;
  delete css.paddingLeft;

  if (options?.preserveBlockTypography) {
    delete css.fontSize;
    delete css.fontWeight;
    delete css.lineHeight;
  }

  return css;
}

/** 执行 `imagesFromImageOnlyParagraph` 封装的DOC 渲染处理步骤。 */
export function imagesFromImageOnlyParagraph(block: DocBlock) {
  // 二进制 DOC 没有稳定的图片锚点模型，这里用“无文字且全是图片”的段落作为图片布局信号。
  if (block.type !== 'paragraph' || block.text.trim()) return [];
  const inlines = block.inlines ?? [];
  const hasVisibleText = inlines.some(
    (inline) => inline.type === 'text' && inline.text.trim(),
  );
  const images = inlines.flatMap((inline) =>
    inline.type === 'image' ? [inline.image] : [],
  );
  if (!images.length || hasVisibleText) return [];
  return images;
}

/** 执行 `canShareImageRow` 封装的DOC 渲染处理步骤。 */
export function canShareImageRow(
  left: DocImage,
  right: DocImage,
  contentWidth: number,
) {
  if (!left.width || !right.width) return false;
  // 只让较小图片并排，避免大图被压缩后影响文档可读性。
  const maxSmallImageWidth = Math.min(300, contentWidth * 0.55);
  return (
    left.width <= maxSmallImageWidth &&
    right.width <= maxSmallImageWidth &&
    left.width + right.width + DOC_IMAGE_ROW_GAP <= contentWidth
  );
}

/** 执行 `imageRows` 封装的DOC 渲染处理步骤。 */
export function imageRows(images: DocImage[], contentWidth: number) {
  const rows: DocImage[][] = [];
  let index = 0;

  while (index < images.length) {
    const current = images[index];
    const next = images[index + 1];
    if (next && canShareImageRow(current, next, contentWidth)) {
      rows.push([current, next]);
      index += 2;
      continue;
    }
    rows.push([current]);
    index += 1;
  }

  return rows;
}

/** 执行 `weightedTextLength` 封装的DOC 渲染处理步骤。 */
function weightedTextLength(text: string) {
  return Array.from(text).reduce(
    (sum, char) => sum + (/[\u4e00-\u9fa5]/.test(char) ? 1 : 0.55),
    0,
  );
}

/** 执行 `estimateLineCount` 封装的DOC 渲染处理步骤。 */
function estimateLineCount(
  text: string,
  width: number,
  fontSize: number,
  minimumCharsPerLine = 8,
) {
  const charsPerLine = Math.max(
    minimumCharsPerLine,
    Math.floor(width / (fontSize * 0.95)),
  );
  return text
    .split('\n')
    .reduce(
      (lines, line) =>
        lines + Math.max(1, Math.ceil(weightedTextLength(line) / charsPerLine)),
      0,
    );
}

/** 估算带左右缩进和首行缩进的 DOC 段落行数。 */
function estimateParagraphLineCount(
  block: DocParagraphBlock,
  contentWidth: number,
  fontSize: number,
) {
  const paragraphWidth = Math.max(
    fontSize * 8,
    contentWidth -
      (block.style?.indentLeft ?? 0) -
      (block.style?.indentRight ?? 0),
  );
  const firstLineWidth = Math.max(
    fontSize * 4,
    paragraphWidth - (block.style?.firstLineIndent ?? 0),
  );
  const normalCapacity = Math.max(
    8,
    Math.floor(paragraphWidth / (fontSize * 0.95)),
  );
  const firstLineCapacity = Math.max(
    4,
    Math.floor(firstLineWidth / (fontSize * 0.95)),
  );

  return (block.text || ' ').split('\n').reduce((lines, line) => {
    const length = weightedTextLength(line);
    if (length <= firstLineCapacity) return lines + 1;
    return lines + 1 + Math.ceil((length - firstLineCapacity) / normalCapacity);
  }, 0);
}

/** 把无单位倍数或已标准化的绝对行距统一换算为像素。 */
function resolveLineHeightPx(fontSize: number, lineHeight: number) {
  return lineHeight > 4 ? lineHeight : fontSize * lineHeight;
}

/** 执行 `estimateParagraphTextHeight` 封装的DOC 渲染处理步骤。 */
function estimateParagraphTextHeight(
  block: DocParagraphBlock,
  contentWidth: number,
) {
  const isTitle = block.role === 'title';
  const isHeading = block.role === 'heading';
  const fontSize =
    block.style?.fontSize ?? (isTitle ? 22 : isHeading ? 16 : 14);
  const lineHeight =
    block.style?.lineHeight ?? (isTitle ? 1.45 : isHeading ? 1.65 : 1.8);
  const defaultSpacingAfter = isTitle ? 18 : isHeading ? 14 : 12;
  const spacingBefore = block.style?.spacingBefore ?? 0;
  const spacingAfter = block.style?.spacingAfter ?? defaultSpacingAfter;
  const padding =
    (block.style?.paddingTop ?? 0) + (block.style?.paddingBottom ?? 0);
  const lines = estimateParagraphLineCount(block, contentWidth, fontSize);

  return Math.max(
    18,
    lines * resolveLineHeightPx(fontSize, lineHeight) +
      spacingBefore +
      spacingAfter +
      padding,
  );
}

/** 按 DOC 图片长宽比和并排数量计算源版式中的渲染宽度。 */
export function getDocImageRenderWidth(
  image: DocImage,
  contentWidth: number,
  rowLength: number,
) {
  const preferredWidth = image.width
    ? Math.min(image.width, contentWidth)
    : contentWidth;
  return rowLength > 1 && image.width
    ? Math.min(image.width, (contentWidth - DOC_IMAGE_ROW_GAP) / rowLength)
    : preferredWidth;
}

/** 执行 `estimateImageHeight` 封装的DOC 渲染处理步骤。 */
function estimateImageHeight(
  image: DocImage,
  contentWidth: number,
  rowLength: number,
) {
  const renderedWidth = getDocImageRenderWidth(image, contentWidth, rowLength);

  if (image.width && image.height) {
    return (image.height * renderedWidth) / image.width;
  }

  return Math.min(240, contentWidth * 0.55);
}

/** 执行 `estimateImageRowHeight` 封装的DOC 渲染处理步骤。 */
function estimateImageRowHeight(row: DocImage[], contentWidth: number) {
  return Math.max(
    ...row.map((image) => estimateImageHeight(image, contentWidth, row.length)),
    0,
  );
}

/** 执行 `estimateTableHeight` 封装的DOC 渲染处理步骤。 */
function estimateTableHeight(
  block: Extract<
    DocBlock,
    {
      /** 用于区分 当前结构 不同结构分支的类型标识。 */ type: 'table';
    }
  >,
  contentWidth: number,
) {
  const columnCount = Math.max(...block.rows.map((row) => row.cells.length), 1);
  const columns = block.columns?.length
    ? block.columns
    : Array.from({ length: columnCount }, () => contentWidth / columnCount);
  const totalColumns =
    columns.reduce((sum, width) => sum + width, 0) || contentWidth;
  const tableContentWidth = block.width ?? contentWidth;

  const rowHeights = block.rows.map((row) => {
    // 空白模板行没有内容参与测量，需要保留可填写的标准行高。
    const minimumRowHeight =
      block.width &&
      row.cells.every((cell) => !cell.text.trim() && !cell.inlines?.length)
        ? 32
        : block.width
        ? 1
        : 28;
    let columnIndex = 0;
    const contentHeight = Math.max(
      minimumRowHeight,
      ...row.cells.map((cell) => {
        const fontSize = cell.style?.fontSize ?? 13;
        const lineHeight = cell.style?.lineHeight ?? 1.65;
        const padding =
          (cell.style?.paddingTop ?? 5) + (cell.style?.paddingBottom ?? 5);
        const horizontalPadding =
          (cell.style?.paddingLeft ?? 8) + (cell.style?.paddingRight ?? 8);
        const colSpan = Math.max(1, cell.colSpan ?? 1);
        const spannedWidth = columns
          .slice(columnIndex, columnIndex + colSpan)
          .reduce((sum, width) => sum + width, 0);
        columnIndex += colSpan;
        const width =
          cell.width ?? (spannedWidth / totalColumns) * tableContentWidth;
        return (
          estimateLineCount(
            cell.text || ' ',
            Math.max(fontSize, width - horizontalPadding),
            fontSize,
            block.width ? 1 : 8,
          ) *
            resolveLineHeightPx(fontSize, lineHeight) +
          padding +
          1
        );
      }),
    );
    if (row.height === undefined) return contentHeight;
    return row.heightRule === 'exact'
      ? row.height
      : Math.max(row.height, contentHeight);
  });

  return (
    rowHeights.reduce((sum, height) => sum + height, 0) +
    (block.spacingAfter ?? 16)
  );
}

/** 执行 `estimateListHeight` 封装的DOC 渲染处理步骤。 */
function estimateListHeight(
  block: Extract<
    DocBlock,
    {
      /** 用于区分 当前结构 不同结构分支的类型标识。 */ type: 'list';
    }
  >,
  contentWidth: number,
) {
  const fontSize = block.style?.fontSize ?? 14;
  const lineHeight = block.style?.lineHeight ?? 1.7;
  const itemSpacingAfter = block.style?.spacingAfter ?? 8;
  const itemHeight = block.items.reduce(
    (sum, item) =>
      sum +
      estimateLineCount(item.text || ' ', contentWidth - 24, fontSize) *
        resolveLineHeightPx(fontSize, lineHeight) +
      itemSpacingAfter,
    0,
  );
  return itemHeight + 16;
}

/** 执行 `estimateBlockHeight` 封装的DOC 渲染处理步骤。 */
function estimateBlockHeight(block: DocBlock, contentWidth: number) {
  if (block.type === 'table') return estimateTableHeight(block, contentWidth);
  if (block.type === 'list') return estimateListHeight(block, contentWidth);

  const images =
    block.inlines?.flatMap((inline) =>
      inline.type === 'image' ? [inline.image] : [],
    ) ?? [];
  const imageHeight = images.reduce(
    (sum, image) => sum + estimateImageHeight(image, contentWidth, 1) + 6,
    0,
  );
  return estimateParagraphTextHeight(block, contentWidth) + imageHeight;
}

/** 估算“与标题同页”所需的下一块首行高度，正文剩余行仍可自然换页。 */
function estimateLeadingBlockHeight(block: DocBlock, contentWidth: number) {
  if (imagesFromImageOnlyParagraph(block).length) {
    return estimateBlockHeight(block, contentWidth);
  }
  if (block.type === 'table') {
    return estimateTableHeight(
      { ...block, rows: block.rows.slice(0, 1) },
      contentWidth,
    );
  }
  if (block.type === 'list') {
    const fontSize = block.style?.fontSize ?? 14;
    const lineHeight = block.style?.lineHeight ?? 1.7;
    return (
      resolveLineHeightPx(fontSize, lineHeight) +
      (block.style?.spacingAfter ?? 8)
    );
  }
  const isTitle = block.role === 'title';
  const isHeading = block.role === 'heading';
  const fontSize =
    block.style?.fontSize ?? (isTitle ? 22 : isHeading ? 16 : 14);
  const lineHeight =
    block.style?.lineHeight ?? (isTitle ? 1.45 : isHeading ? 1.65 : 1.8);
  return (
    resolveLineHeightPx(fontSize, lineHeight) +
    (block.style?.spacingBefore ?? 0) +
    (block.style?.paddingTop ?? 0)
  );
}

/** 创建 `createImageParagraphBlock` 返回的对象，供DOC 渲染使用。 */
function createImageParagraphBlock(
  id: string,
  images: DocImage[],
  style?: DocTextStyle,
): DocParagraphBlock {
  const inlines: DocTextInline[] = images.map((image) => ({
    type: 'image',
    image,
  }));
  return {
    id,
    type: 'paragraph',
    text: '',
    inlines,
    role: 'body',
    style,
  };
}

function isUnbrokenDocHeading(block: DocBlock) {
  return (
    block.type === 'paragraph' &&
    block.role === 'heading' &&
    !block.pageBreakBefore
  );
}

/** 找出必须留到下一批才能判断“标题跟图”和连续图片的尾部起点。 */
function findDocPaginationCarryStart(blocks: readonly DocBlock[]) {
  let index = blocks.length;
  while (index > 0 && imagesFromImageOnlyParagraph(blocks[index - 1]).length) {
    index -= 1;
  }
  if (index < blocks.length) {
    while (index > 0 && isUnbrokenDocHeading(blocks[index - 1])) {
      index -= 1;
    }
    return index;
  }
  while (index > 0 && isUnbrokenDocHeading(blocks[index - 1])) {
    index -= 1;
  }
  return index < blocks.length ? index : blocks.length;
}

/** DOC 同步与异步分页共用的增量状态机。 */
export class DocPaginationState {
  private readonly contentHeight: number;
  private readonly readyPages: PaginatedDocPage[] = [];
  private currentBlocks: DocBlock[] = [];
  private carryBlocks: DocBlock[] = [];
  private currentHeight = 0;
  private syntheticImageIndex = 0;
  private pageSequence = 0;
  private completed = false;
  private sawBlock = false;

  constructor(page: DocPage, private readonly contentWidth: number) {
    this.contentHeight = Math.max(
      240,
      page.minHeight -
        page.marginTop -
        page.marginBottom -
        DOC_PAGE_HEIGHT_BUFFER +
        DOC_PAGE_HEIGHT_TOLERANCE,
    );
  }

  append(blocks: readonly DocBlock[], complete = false) {
    if (this.completed) throw new Error('DOC 分页状态已经完成');
    this.sawBlock ||= blocks.length > 0;
    const pending = [...this.carryBlocks, ...blocks];
    const processEnd = complete
      ? pending.length
      : findDocPaginationCarryStart(pending);
    this.carryBlocks = pending.slice(processEnd);
    this.processBlocks(pending.slice(0, processEnd));
    if (complete) {
      this.processBlocks(this.carryBlocks);
      this.carryBlocks = [];
      this.flushPage();
      if (!this.sawBlock && !this.readyPages.length) {
        this.readyPages.push({ id: 'doc-page-1', blocks: [] });
      }
      this.completed = true;
    }
    return this.drain();
  }

  private drain() {
    return this.readyPages.splice(0);
  }

  private flushPage() {
    if (!this.currentBlocks.length) return;
    const looksLikeCover =
      this.currentBlocks.some(
        (block) => block.type === 'paragraph' && block.role === 'title',
      ) &&
      this.currentBlocks.some(
        (block) => imagesFromImageOnlyParagraph(block).length,
      ) &&
      !this.currentBlocks.some((block) => block.type === 'table');
    const pageBlocks = looksLikeCover
      ? this.currentBlocks.map((block) =>
          block.type === 'paragraph' && block.role === 'title'
            ? {
                ...block,
                style: {
                  ...block.style,
                  spacingBefore: Math.max(197, block.style?.spacingBefore ?? 0),
                  spacingAfter: 4,
                },
              }
            : block,
        )
      : this.currentBlocks;
    this.pageSequence += 1;
    // 图片封面依靠大量空段落垂直居中，空段清理后用等价段前距恢复位置。
    this.readyPages.push({
      id: `doc-page-${this.pageSequence}`,
      blocks: pageBlocks,
    });
    this.currentBlocks = [];
    this.currentHeight = 0;
  }

  private appendBlock(block: DocBlock, estimatedHeight: number) {
    if (
      this.currentBlocks.length &&
      this.currentHeight + estimatedHeight > this.contentHeight
    ) {
      this.flushPage();
    }
    this.currentBlocks.push(block);
    this.currentHeight += estimatedHeight;
  }

  private appendImageRows(images: DocImage[], style?: DocTextStyle) {
    const rows = imageRows(images, this.contentWidth);
    let pendingImages: DocImage[] = [];
    let pendingHeight = DOC_IMAGE_LAYOUT_VERTICAL_MARGIN;
    const flushImages = () => {
      if (!pendingImages.length) return;
      this.syntheticImageIndex += 1;
      this.appendBlock(
        createImageParagraphBlock(
          `doc-image-page-group-${this.syntheticImageIndex}`,
          pendingImages,
          style,
        ),
        pendingHeight,
      );
      pendingImages = [];
      pendingHeight = DOC_IMAGE_LAYOUT_VERTICAL_MARGIN;
    };
    rows.forEach((row) => {
      const rowHeight =
        estimateImageRowHeight(row, this.contentWidth) +
        (pendingImages.length ? DOC_IMAGE_LAYOUT_ROW_GAP : 0);
      if (
        pendingImages.length &&
        this.currentHeight + pendingHeight + rowHeight > this.contentHeight
      ) {
        flushImages();
      }
      pendingImages.push(...row);
      pendingHeight += rowHeight;
    });
    flushImages();
  }

  private processBlocks(blocks: readonly DocBlock[]) {
    let index = 0;
    while (index < blocks.length) {
      const currentBlock = blocks[index];
      if (currentBlock.type === 'paragraph' && currentBlock.pageBreakBefore) {
        this.flushPage();
        index += 1;
        continue;
      }
      if (
        currentBlock.type === 'paragraph' &&
        currentBlock.role === 'heading'
      ) {
        let keepIndex = index;
        let keepHeight = 0;
        while (keepIndex < blocks.length) {
          const keepBlock = blocks[keepIndex];
          if (
            keepBlock.type !== 'paragraph' ||
            keepBlock.role !== 'heading' ||
            keepBlock.pageBreakBefore
          ) {
            break;
          }
          keepHeight += estimateBlockHeight(keepBlock, this.contentWidth);
          keepIndex += 1;
        }
        const followingBlock = blocks[keepIndex];
        if (followingBlock) {
          keepHeight += estimateLeadingBlockHeight(
            followingBlock,
            this.contentWidth,
          );
        }
        if (
          followingBlock &&
          this.currentBlocks.length &&
          this.currentHeight + keepHeight > this.contentHeight
        ) {
          this.flushPage();
        }
      }
      const currentPageLooksLikeCover =
        this.currentBlocks.some(
          (block) => block.type === 'paragraph' && block.role === 'title',
        ) &&
        this.currentBlocks.some(
          (block) => imagesFromImageOnlyParagraph(block).length,
        );
      if (currentBlock.type === 'table' && currentPageLooksLikeCover) {
        this.flushPage();
      }
      const currentTableHeight =
        currentBlock.type === 'table'
          ? estimateTableHeight(currentBlock, this.contentWidth)
          : 0;
      if (
        currentBlock.type === 'table' &&
        (currentTableHeight > this.contentHeight ||
          (this.currentBlocks.length &&
            this.currentHeight + currentTableHeight > this.contentHeight))
      ) {
        this.appendTableParts(currentBlock);
        index += 1;
        continue;
      }
      const imageOnlyParagraphImages =
        imagesFromImageOnlyParagraph(currentBlock);
      if (imageOnlyParagraphImages.length) {
        const imageGroup = [...imageOnlyParagraphImages];
        const imageStyle =
          currentBlock.type === 'paragraph' ? currentBlock.style : undefined;
        const imageAlignment = imageStyle?.textAlign;
        let nextIndex = index + 1;
        while (nextIndex < blocks.length) {
          const nextBlock = blocks[nextIndex];
          const nextImages = imagesFromImageOnlyParagraph(nextBlock);
          if (!nextImages.length) break;
          const nextAlignment =
            nextBlock.type === 'paragraph'
              ? nextBlock.style?.textAlign
              : undefined;
          if (nextAlignment !== imageAlignment) break;
          imageGroup.push(...nextImages);
          nextIndex += 1;
        }
        this.appendImageRows(imageGroup, imageStyle);
        index = nextIndex;
        continue;
      }
      this.appendBlock(
        currentBlock,
        estimateBlockHeight(currentBlock, this.contentWidth),
      );
      index += 1;
    }
  }

  private appendTableParts(currentBlock: Extract<DocBlock, { type: 'table' }>) {
    let rowIndex = 0;
    let partIndex = 0;
    while (rowIndex < currentBlock.rows.length) {
      let availableHeight = this.contentHeight - this.currentHeight;
      if (this.currentBlocks.length && availableHeight < 80) {
        this.flushPage();
        availableHeight = this.contentHeight;
      }
      const startRowIndex = rowIndex;
      const spacingAfter = currentBlock.spacingAfter ?? 16;
      let rowsHeight = 0;
      while (rowIndex < currentBlock.rows.length) {
        const row = currentBlock.rows[rowIndex];
        const rowHeight =
          estimateTableHeight(
            { ...currentBlock, rows: [row] },
            this.contentWidth,
          ) - spacingAfter;
        if (
          rowIndex > startRowIndex &&
          rowsHeight + rowHeight > availableHeight
        ) {
          break;
        }
        rowsHeight += rowHeight;
        rowIndex += 1;
      }
      partIndex += 1;
      const isFinalPart = rowIndex >= currentBlock.rows.length;
      const partSpacingAfter = isFinalPart
        ? Math.min(spacingAfter, Math.max(0, availableHeight - rowsHeight))
        : 0;
      this.appendBlock(
        {
          ...currentBlock,
          id: `${currentBlock.id}-part-${partIndex}`,
          sourceBlockId: currentBlock.sourceBlockId ?? currentBlock.id,
          rows: currentBlock.rows.slice(startRowIndex, rowIndex),
          spacingAfter: partSpacingAfter,
        },
        rowsHeight + partSpacingAfter,
      );
      if (rowIndex < currentBlock.rows.length) this.flushPage();
    }
  }
}

/**
 * 按 DOC 页面可用高度估算分页，并把连续图片段落拆到多个页面中。
 * 同步入口与渐进 Source 使用同一个状态机，保证页 ID 和块顺序一致。
 */
export function paginateDocBlocks(
  blocks: DocBlock[],
  page: DocPage,
  contentWidth: number,
): PaginatedDocPage[] {
  return new DocPaginationState(page, contentWidth).append(blocks, true);
}
