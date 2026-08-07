// DOC 二进制格式缺少显式页面模型，分页服务统一估算页面块与图片布局。
import type {
  DocBlock,
  DocImage,
  DocPage,
  DocParagraphBlock,
  DocTextInline,
  DocTextStyle,
} from './types';

/** DOC 图片并排布局时的行内间距，单位为标准化渲染像素。 */
export const DOC_IMAGE_ROW_GAP = 6;

// 页面可用高度已扣除上下页边距；不再重复预留缓冲，避免把本应同页的图片组拆页。
/** DOC 分页高度估算预留量，单位为标准化渲染像素。 */
const DOC_PAGE_HEIGHT_BUFFER = 0;
// 浏览器行盒会累计小数像素，允许极小误差可避免 1px 级伪溢出触发整页后移。
/** DOC 分页高度比较允许的误差，单位为标准化渲染像素。 */
const DOC_PAGE_HEIGHT_TOLERANCE = 2;
/** DOC 图片布局相邻行之间的间距，单位为标准化渲染像素。 */
const DOC_IMAGE_LAYOUT_ROW_GAP = 12;
/** DOC 图片布局上下预留的空白，单位为标准化渲染像素。 */
const DOC_IMAGE_LAYOUT_VERTICAL_MARGIN = 10;

// DOC 没有 OOXML 的显式页面模型，这里记录前端估算分页后的页面块集合。
/** 完成内容块分页后的 DOC 页面。 */
export type PaginatedDocPage = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 按源文档顺序排列的内容块。 */
  blocks: DocBlock[];
};

/** 收集内容块中需要按页面坐标叠放的绘图画布。 */
export function pageDrawingImagesFromBlock(block: DocBlock) {
  const inlines =
    block.type === 'paragraph'
      ? block.inlines ?? []
      : block.type === 'table'
      ? block.rows.flatMap((row) =>
          row.cells.flatMap((cell) => cell.inlines ?? []),
        )
      : block.items.flatMap((item) => item.inlines ?? []);
  return inlines.flatMap((inline) =>
    inline.type === 'image' && inline.image.pageDrawingLayer
      ? [inline.image]
      : [],
  );
}

/** 判断段落是否只承担页级绘图锚点，不应消耗正文排版高度。 */
export function isPageDrawingOnlyParagraph(block: DocBlock) {
  if (block.type !== 'paragraph' || block.text.trim()) return false;
  const visibleInlines = (block.inlines ?? []).filter(
    (inline) => inline.type === 'image' || inline.text.trim(),
  );
  return (
    visibleInlines.length > 0 &&
    visibleInlines.every(
      (inline) => inline.type === 'image' && inline.image.pageDrawingLayer,
    )
  );
}

/** 提取仅包含图片的 DOC 段落资源。 */
export function imagesFromImageOnlyParagraph(block: DocBlock) {
  // 二进制 DOC 没有稳定的图片锚点模型，这里用“无文字且全是图片”的段落作为图片布局信号。
  if (block.type !== 'paragraph' || block.text.trim()) return [];
  // 混合页级画布的段落交给普通块渲染，避免图片组重建时丢失页面锚点。
  if (pageDrawingImagesFromBlock(block).length) return [];
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

/** 判断相邻图片能否在同一行排版。 */
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

/** 按可用内容宽度将 DOC 图片分组为多行。 */
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

function weightedTextLength(text: string) {
  return Array.from(text).reduce(
    (sum, char) => sum + (/[\u4e00-\u9fa5]/.test(char) ? 1 : 0.55),
    0,
  );
}

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
  if (image.pageInsets && image.width) return image.width;
  const preferredWidth = image.width
    ? Math.min(image.width, contentWidth)
    : contentWidth;
  return rowLength > 1 && image.width
    ? Math.min(image.width, (contentWidth - DOC_IMAGE_ROW_GAP) / rowLength)
    : preferredWidth;
}

function estimateImageHeight(
  image: DocImage,
  contentWidth: number,
  rowLength: number,
) {
  const renderedWidth = getDocImageRenderWidth(image, contentWidth, rowLength);

  if (image.width && image.height) {
    const renderedHeight = (image.height * renderedWidth) / image.width;
    return image.pageInsets
      ? Math.max(
          0,
          renderedHeight - image.pageInsets.top - image.pageInsets.bottom,
        )
      : renderedHeight;
  }

  return Math.min(240, contentWidth * 0.55);
}

function estimateImageRowHeight(row: DocImage[], contentWidth: number) {
  return Math.max(
    ...row.map((image) => estimateImageHeight(image, contentWidth, row.length)),
    0,
  );
}

function estimateTableHeight(
  block: Extract<
    DocBlock,
    {
      /** 固定为 `table`，用于区分联合类型分支。 */
      type: 'table';
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
    (block.spacingBefore ?? 0) +
    rowHeights.reduce((sum, height) => sum + height, 0) +
    (block.spacingAfter ?? 16)
  );
}

function estimateListHeight(
  block: Extract<
    DocBlock,
    {
      /** 固定为 `list`，用于区分联合类型分支。 */
      type: 'list';
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

function estimateBlockHeight(block: DocBlock, contentWidth: number) {
  if (block.type === 'table') return estimateTableHeight(block, contentWidth);
  if (block.type === 'list') return estimateListHeight(block, contentWidth);

  const images =
    block.inlines?.flatMap((inline) =>
      inline.type === 'image' && !inline.image.pageDrawingLayer
        ? [inline.image]
        : [],
    ) ?? [];
  const imageHeight = images.reduce(
    (sum, image) => sum + estimateImageHeight(image, contentWidth, 1) + 6,
    0,
  );
  return estimateParagraphTextHeight(block, contentWidth) + imageHeight;
}

/** 读取块前垂直间距，供分页时把连续空段正确消耗在当前页尾。 */
function blockSpacingBefore(block: DocBlock) {
  return block.type === 'table'
    ? block.spacingBefore ?? 0
    : block.style?.spacingBefore ?? 0;
}

/** 返回仅调整块前间距的副本，避免分页修正污染解析得到的源模型。 */
function withBlockSpacingBefore(block: DocBlock, spacingBefore: number) {
  if (block.type === 'table') {
    return {
      ...block,
      spacingBefore: spacingBefore || undefined,
    };
  }
  return {
    ...block,
    style: {
      ...block.style,
      spacingBefore: spacingBefore || undefined,
    },
  };
}

/** 估算“与标题同页”所需的下一块首行高度，正文剩余行仍可自然换页。 */
function estimateLeadingBlockHeight(block: DocBlock, contentWidth: number) {
  if (imagesFromImageOnlyParagraph(block).length) {
    return estimateBlockHeight(block, contentWidth);
  }
  if (block.type === 'table') {
    // 表头必须与首个数据行同页，避免页尾只留下标题或孤立表头。
    return estimateTableHeight(
      { ...block, rows: block.rows.slice(0, 2) },
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
    if (images.length === 1 && images[0].pageInsets) {
      this.syntheticImageIndex += 1;
      this.appendBlock(
        createImageParagraphBlock(
          `doc-image-page-group-${this.syntheticImageIndex}`,
          images,
          style,
        ),
        estimateImageHeight(images[0], this.contentWidth, 1),
      );
      return;
    }
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
      let currentBlock = blocks[index];
      const pageDrawingOnly = isPageDrawingOnlyParagraph(currentBlock);
      const leadingSpacing = pageDrawingOnly
        ? 0
        : blockSpacingBefore(currentBlock);
      if (leadingSpacing > 0 && this.currentBlocks.length) {
        const availableHeight = Math.max(
          0,
          this.contentHeight - this.currentHeight,
        );
        if (leadingSpacing >= availableHeight) {
          // 解析器把连续空段折算为下一块的段前距；跨页时先消耗当前页剩余空间，
          // Word 不会把段前距重复带到新页页首，清零后可避免空白逐页累积。
          this.flushPage();
          currentBlock = withBlockSpacingBefore(currentBlock, 0);
        }
      }
      if (currentBlock.type === 'paragraph' && currentBlock.pageBreakBefore) {
        this.flushPage();
        if (pageDrawingOnly) this.appendBlock(currentBlock, 0);
        index += 1;
        continue;
      }
      if (pageDrawingOnly) {
        this.appendBlock(currentBlock, 0);
        index += 1;
        continue;
      }
      const followingTable = blocks[index + 1];
      const isTableCaption =
        currentBlock.type === 'paragraph' &&
        !currentBlock.pageBreakBefore &&
        currentBlock.text.replace(/\s+/g, '').length <= 40 &&
        followingTable?.type === 'table';
      if (isTableCaption) {
        // 单页可容纳的表格与表题整体保持；超长表格至少保留表头和首个数据行。
        const fullTableHeight = estimateBlockHeight(
          followingTable,
          this.contentWidth,
        );
        const keepHeight =
          estimateBlockHeight(currentBlock, this.contentWidth) +
          (fullTableHeight <= this.contentHeight
            ? fullTableHeight
            : estimateLeadingBlockHeight(followingTable, this.contentWidth));
        if (
          this.currentBlocks.length &&
          this.currentHeight + keepHeight > this.contentHeight
        ) {
          this.flushPage();
        }
      } else if (
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
      const partSpacingBefore =
        partIndex === 0 ? currentBlock.spacingBefore ?? 0 : 0;
      let availableHeight = this.contentHeight - this.currentHeight;
      if (this.currentBlocks.length && availableHeight < 80) {
        this.flushPage();
        availableHeight = this.contentHeight;
      }
      availableHeight = Math.max(0, availableHeight - partSpacingBefore);
      const startRowIndex = rowIndex;
      const spacingAfter = currentBlock.spacingAfter ?? 16;
      let rowsHeight = 0;
      while (rowIndex < currentBlock.rows.length) {
        const row = currentBlock.rows[rowIndex];
        const rowHeight = estimateTableHeight(
          {
            ...currentBlock,
            rows: [row],
            spacingBefore: 0,
            spacingAfter: 0,
          },
          this.contentWidth,
        );
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
          spacingBefore: partSpacingBefore || undefined,
          spacingAfter: partSpacingAfter,
        },
        partSpacingBefore + rowsHeight + partSpacingAfter,
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
