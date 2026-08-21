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
/** 结构化 DOC 表格相对正文网格的顶部修正量。 */
export const DOC_STRUCTURED_TABLE_TOP_OFFSET = 8;

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
/** 表格至少可在页尾显示此比例时才拆行，避免只留下表头和孤立首行。 */
const DOC_TABLE_SPLIT_MIN_VISIBLE_RATIO = 0.5;
/** 常用汉字在中文字体中按全角字宽参与换行估算。 */
const DOC_HAN_CHAR_PATTERN = /[\u4e00-\u9fa5]/;
/** 12pt 及以上中文正文中与汉字占用相同网格的全角字符范围。 */
const DOC_FULL_WIDTH_CHAR_PATTERN =
  /[\u2e80-\u9fff\uf900-\ufaff\uff01-\uff60\uffe0-\uffe6]/;
/** 12pt 中文正文换算成标准化渲染像素后的字号。 */
const DOC_FULL_WIDTH_PUNCTUATION_FONT_SIZE = 16;

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
    (inline) =>
      inline.type === 'image' ||
      (inline.type === 'text' && Boolean(inline.text.trim())),
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

function weightedTextLength(text: string, fontSize: number) {
  const fullWidthPattern =
    fontSize >= DOC_FULL_WIDTH_PUNCTUATION_FONT_SIZE
      ? DOC_FULL_WIDTH_CHAR_PATTERN
      : DOC_HAN_CHAR_PATTERN;
  return Array.from(text).reduce((sum, char) => {
    // 12pt 中文网格中的全角标点占满字格；较小字号沿用 Word 的紧凑标点宽度。
    if (char === '\t') return sum + 4;
    return sum + (fullWidthPattern.test(char) ? 1 : 0.55);
  }, 0);
}

/** 读取浏览器实际渲染的行内文字，保留编号制表符等分页所需字符。 */
function renderedTextFromInlines(
  fallbackText: string,
  inlines?: DocTextInline[],
) {
  const inlineText = inlines
    ?.flatMap((inline) => (inline.type === 'text' ? [inline.text] : []))
    .join('');
  return inlineText?.length ? inlineText : fallbackText;
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
        lines +
        Math.max(
          1,
          Math.ceil(weightedTextLength(line, fontSize) / charsPerLine),
        ),
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
  // 12pt 两端对齐正文按完整中文网格换行；较小字号与非两端对齐正文保留紧凑字宽。
  const characterWidthFactor =
    fontSize >= DOC_FULL_WIDTH_PUNCTUATION_FONT_SIZE &&
    block.style?.textAlign === 'justify'
      ? 1
      : 0.95;
  const normalCapacity = Math.max(
    8,
    Math.floor(paragraphWidth / (fontSize * characterWidthFactor)),
  );
  const firstLineCapacity = Math.max(
    4,
    Math.floor(firstLineWidth / (fontSize * characterWidthFactor)),
  );
  const wrappedLineCapacity = normalCapacity;

  return renderedTextFromInlines(block.text, block.inlines)
    .split('\n')
    .reduce((lines, line) => {
      const length = weightedTextLength(line, fontSize);
      if (length <= firstLineCapacity) return lines + 1;
      return (
        lines +
        1 +
        Math.ceil((length - firstLineCapacity) / wrappedLineCapacity)
      );
    }, 0);
}

/** 读取段落及其文字片段实际使用的最大字号，避免行内样式放大后分页仍按段落默认值估算。 */
function resolveParagraphFontSize(
  block: DocParagraphBlock,
  fallbackFontSize: number,
) {
  return Math.max(
    block.style?.fontSize ?? fallbackFontSize,
    ...(block.inlines ?? []).flatMap((inline) =>
      inline.type === 'text' && inline.style?.fontSize !== undefined
        ? [inline.style.fontSize]
        : [],
    ),
  );
}

/** 把无单位倍数或已标准化的绝对行距统一换算为像素。 */
function resolveLineHeightPx(fontSize: number, lineHeight: number) {
  return lineHeight > 4 ? lineHeight : fontSize * lineHeight;
}

type DocParagraphLayoutMetrics = {
  fontSize: number;
  lineHeightPx: number;
  spacingBefore: number;
  spacingAfter: number;
  paddingTop: number;
  paddingBottom: number;
};

/** 统一读取 DOC 段落估高与跨页拆分共用的垂直排版参数。 */
function resolveParagraphLayoutMetrics(
  block: DocParagraphBlock,
): DocParagraphLayoutMetrics {
  const isTitle = block.role === 'title';
  const isHeading = block.role === 'heading';
  const fontSize = resolveParagraphFontSize(
    block,
    isTitle ? 22 : isHeading ? 16 : 14,
  );
  const lineHeight =
    block.style?.lineHeight ?? (isTitle ? 1.45 : isHeading ? 1.65 : 1.8);
  return {
    fontSize,
    lineHeightPx: resolveLineHeightPx(fontSize, lineHeight),
    spacingBefore: block.style?.spacingBefore ?? 0,
    spacingAfter:
      block.style?.spacingAfter ?? (isTitle ? 18 : isHeading ? 14 : 12),
    paddingTop: block.style?.paddingTop ?? 0,
    paddingBottom: block.style?.paddingBottom ?? 0,
  };
}

function estimateParagraphTextHeight(
  block: DocParagraphBlock,
  contentWidth: number,
) {
  const metrics = resolveParagraphLayoutMetrics(block);
  const lines = estimateParagraphLineCount(
    block,
    contentWidth,
    metrics.fontSize,
  );
  return Math.max(
    18,
    lines * metrics.lineHeightPx +
      metrics.spacingBefore +
      metrics.spacingAfter +
      metrics.paddingTop +
      metrics.paddingBottom,
  );
}
/** 按可见文字偏移拆分行内节点，同时保留两侧文字样式和超链接。 */
function splitTextInlinesAtOffset(
  inlines: DocTextInline[] | undefined,
  offset: number,
) {
  const leading: DocTextInline[] = [];
  const trailing: DocTextInline[] = [];
  let remaining = offset;
  let reachedTrailing = false;
  (inlines ?? []).forEach((inline) => {
    if (inline.type !== 'text') {
      (reachedTrailing ? trailing : leading).push(inline);
      return;
    }
    if (reachedTrailing || remaining <= 0) {
      reachedTrailing = true;
      trailing.push(inline);
      return;
    }
    if (inline.text.length <= remaining) {
      leading.push(inline);
      remaining -= inline.text.length;
      return;
    }
    leading.push({ ...inline, text: inline.text.slice(0, remaining) });
    trailing.push({ ...inline, text: inline.text.slice(remaining) });
    remaining = 0;
    reachedTrailing = true;
  });
  return { leading, trailing };
}

/** 根据真实行数比例定位分页文字边界，兼顾中英文混排的近似字宽。 */
function resolveParagraphSplitOffset(
  text: string,
  fontSize: number,
  leadingLines: number,
  totalLines: number,
) {
  const totalWeight = weightedTextLength(text, fontSize);
  const targetWeight = (totalWeight * leadingLines) / totalLines;
  let consumedWeight = 0;
  let offset = 0;
  for (const char of Array.from(text)) {
    const nextWeight = weightedTextLength(char, fontSize);
    if (consumedWeight + nextWeight > targetWeight && offset > 0) break;
    consumedWeight += nextWeight;
    offset += char.length;
  }
  return Math.min(Math.max(1, offset), Math.max(1, text.length - 1));
}
/** 把允许跨页的正文段落拆成当前页片段和后续页片段。 */
function splitParagraphForAvailableHeight(
  block: DocParagraphBlock,
  totalHeight: number,
  availableHeight: number,
) {
  if (
    block.role === 'title' ||
    block.role === 'heading' ||
    block.isTableOfContents ||
    block.inlines?.some((inline) => inline.type === 'image')
  ) {
    return undefined;
  }
  const text = renderedTextFromInlines(block.text, block.inlines);
  if (!text.trim()) return undefined;
  const metrics = resolveParagraphLayoutMetrics(block);
  const fixedHeight =
    metrics.spacingBefore +
    metrics.spacingAfter +
    metrics.paddingTop +
    metrics.paddingBottom;
  const totalLines = Math.max(
    1,
    Math.round((totalHeight - fixedHeight) / metrics.lineHeightPx),
  );
  const leadingLines = Math.min(
    totalLines - 1,
    Math.floor(
      (availableHeight - metrics.spacingBefore - metrics.paddingTop) /
        metrics.lineHeightPx,
    ),
  );
  if (leadingLines < 1) return undefined;

  const offset = resolveParagraphSplitOffset(
    text,
    metrics.fontSize,
    leadingLines,
    totalLines,
  );
  if (offset <= 0 || offset >= text.length) return undefined;
  const { leading, trailing } = splitTextInlinesAtOffset(block.inlines, offset);
  const sourceBlockId = block.sourceBlockId ?? block.id;
  const leadingBlock: DocParagraphBlock = {
    ...block,
    id: `${block.id}-page-head`,
    sourceBlockId,
    text: text.slice(0, offset),
    inlines: leading.length ? leading : undefined,
    estimatedHeight: undefined,
    style: {
      ...block.style,
      spacingAfter: 0,
      paddingBottom: 0,
    },
  };
  const trailingBlock: DocParagraphBlock = {
    ...block,
    id: `${block.id}-page-tail`,
    sourceBlockId,
    text: text.slice(offset),
    inlines: trailing.length ? trailing : undefined,
    role: 'body',
    outlineLevel: undefined,
    isTableOfContents: undefined,
    pageBreakBefore: undefined,
    estimatedHeight: undefined,
    style: {
      ...block.style,
      firstLineIndent: 0,
      spacingBefore: 0,
      paddingTop: 0,
    },
  };
  return {
    leadingBlock,
    leadingHeight:
      metrics.spacingBefore +
      metrics.paddingTop +
      leadingLines * metrics.lineHeightPx,
    trailingBlock,
    trailingHeight:
      (totalLines - leadingLines) * metrics.lineHeightPx +
      metrics.spacingAfter +
      metrics.paddingBottom,
  };
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
  includeTopOffset = true,
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
            renderedTextFromInlines(cell.text, cell.inlines) || ' ',
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
    (includeTopOffset && block.width ? DOC_STRUCTURED_TABLE_TOP_OFFSET : 0) +
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
  const indentLeft = Math.max(20, block.style?.indentLeft ?? 0);
  const listContentWidth = Math.max(
    fontSize * 4,
    contentWidth -
      indentLeft -
      (block.style?.indentRight ?? 0) +
      fontSize * 0.5,
  );
  const itemHeight = block.items.reduce((sum, item) => {
    const lines = estimateLineCount(
      renderedTextFromInlines(item.text, item.inlines) || ' ',
      listContentWidth,
      fontSize,
    );
    return (
      sum + lines * resolveLineHeightPx(fontSize, lineHeight) + itemSpacingAfter
    );
  }, 0);
  return (
    itemHeight +
    (!block.continuesOnNext && block.style?.spacingAfter === undefined ? 16 : 0)
  );
}

/** 估算列表单项高度，供列表在页尾按项目边界拆分。 */
function estimateListItemHeight(
  block: Extract<DocBlock, { type: 'list' }>,
  item: Extract<DocBlock, { type: 'list' }>['items'][number],
  contentWidth: number,
) {
  return estimateListHeight(
    {
      ...block,
      items: [item],
      continuesOnNext: true,
    },
    contentWidth,
  );
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
  const fontSize = resolveParagraphFontSize(
    block,
    isTitle ? 22 : isHeading ? 16 : 14,
  );
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

  constructor(
    page: DocPage,
    private readonly contentWidth: number,
    private readonly measuredBlockHeights?: ReadonlyMap<string, number>,
  ) {
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
    const sourceBlockId = block.sourceBlockId ?? block.id;
    const measuredHeight =
      block.id === sourceBlockId
        ? this.measuredBlockHeights?.get(sourceBlockId)
        : undefined;
    const calibratedHeight = Math.max(estimatedHeight, measuredHeight ?? 0);
    const availableHeight = this.contentHeight - this.currentHeight;
    const canSplitMeasuredParagraph =
      block.type === 'paragraph' &&
      measuredHeight !== undefined &&
      measuredHeight > estimatedHeight + 0.75;
    const canSplitOrdinaryParagraph =
      block.type === 'paragraph' &&
      measuredHeight === undefined &&
      estimatedHeight > availableHeight;
    if (
      (canSplitMeasuredParagraph && calibratedHeight > availableHeight) ||
      canSplitOrdinaryParagraph
    ) {
      const split = splitParagraphForAvailableHeight(
        block,
        calibratedHeight,
        availableHeight,
      );
      if (split) {
        this.appendBlock(split.leadingBlock, split.leadingHeight);
        this.flushPage();
        this.appendBlock(split.trailingBlock, split.trailingHeight);
        return;
      }
    }
    if (
      this.currentBlocks.length &&
      this.currentHeight + calibratedHeight > this.contentHeight
    ) {
      this.flushPage();
    }
    const measuredBlock =
      block.type === 'table' ? block : { ...block, estimatedHeight };
    this.currentBlocks.push(measuredBlock);
    this.currentHeight += calibratedHeight;
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
        if (
          !currentBlock.text.trim() &&
          !(currentBlock.inlines ?? []).some(
            (inline) => inline.type === 'image',
          )
        ) {
          // 纯分页符只负责结束上一页，本身不应在新页额外生成一个空行盒。
          index += 1;
          continue;
        }
        if (pageDrawingOnly) {
          this.appendBlock(currentBlock, 0);
          index += 1;
          continue;
        }
        // pageBreakBefore 只决定段落从新页开始，段落正文仍必须进入分页结果。
        currentBlock = { ...currentBlock, pageBreakBefore: undefined };
      }
      if (pageDrawingOnly) {
        this.appendBlock(currentBlock, 0);
        index += 1;
        continue;
      }
      const followingFlowBlock = blocks[index + 1];
      const isFlowCaption =
        currentBlock.type === 'paragraph' &&
        !currentBlock.pageBreakBefore &&
        currentBlock.text.trim().length > 0 &&
        !(currentBlock.inlines ?? []).some(
          (inline) => inline.type === 'image',
        ) &&
        currentBlock.text.replace(/\s+/g, '').length <= 40 &&
        (followingFlowBlock?.type === 'table' ||
          followingFlowBlock?.type === 'list');
      if (isFlowCaption && followingFlowBlock) {
        const captionHeight = estimateBlockHeight(
          currentBlock,
          this.contentWidth,
        );
        let followingKeepHeight = estimateLeadingBlockHeight(
          followingFlowBlock,
          this.contentWidth,
        );
        if (followingFlowBlock.type === 'table') {
          const fullTableHeight = estimateBlockHeight(
            followingFlowBlock,
            this.contentWidth,
          );
          const availableTableHeight = Math.max(
            0,
            this.contentHeight - this.currentHeight - captionHeight,
          );
          const canUseCurrentPageForTable =
            fullTableHeight > this.contentHeight ||
            availableTableHeight / Math.max(1, fullTableHeight) >=
              DOC_TABLE_SPLIT_MIN_VISIBLE_RATIO;
          // 页尾只能留下表头和孤立首行时，把可单页容纳的整表移至下一页。
          if (
            fullTableHeight <= this.contentHeight &&
            !canUseCurrentPageForTable
          ) {
            followingKeepHeight = fullTableHeight;
          }
        }
        const keepHeight = captionHeight + followingKeepHeight;
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
      const currentListHeight =
        currentBlock.type === 'list'
          ? estimateListHeight(currentBlock, this.contentWidth)
          : 0;
      if (
        currentBlock.type === 'list' &&
        (currentListHeight > this.contentHeight ||
          (this.currentBlocks.length > 0 &&
            this.currentHeight + currentListHeight > this.contentHeight))
      ) {
        this.appendListParts(currentBlock);
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

  /** 在列表项边界拆分页尾内容，避免整组列表被推到下一页。 */
  private appendListParts(currentBlock: Extract<DocBlock, { type: 'list' }>) {
    let itemIndex = 0;
    let partIndex = 0;
    while (itemIndex < currentBlock.items.length) {
      let availableHeight = this.contentHeight - this.currentHeight;
      const firstItemHeight = estimateListItemHeight(
        currentBlock,
        currentBlock.items[itemIndex],
        this.contentWidth,
      );
      if (this.currentBlocks.length && firstItemHeight > availableHeight) {
        this.flushPage();
        availableHeight = this.contentHeight;
      }
      const startItemIndex = itemIndex;
      let itemsHeight = 0;
      while (itemIndex < currentBlock.items.length) {
        const itemHeight = estimateListItemHeight(
          currentBlock,
          currentBlock.items[itemIndex],
          this.contentWidth,
        );
        const candidateFinal = itemIndex + 1 >= currentBlock.items.length;
        const outerSpacing =
          candidateFinal && currentBlock.style?.spacingAfter === undefined
            ? 16
            : 0;
        if (
          itemIndex > startItemIndex &&
          itemsHeight + itemHeight + outerSpacing > availableHeight
        ) {
          break;
        }
        itemsHeight += itemHeight;
        itemIndex += 1;
        if (itemsHeight + outerSpacing > availableHeight) break;
      }
      if (itemIndex === startItemIndex) {
        itemsHeight = firstItemHeight;
        itemIndex += 1;
      }
      partIndex += 1;
      const isFinalPart = itemIndex >= currentBlock.items.length;
      const partOuterSpacing =
        isFinalPart && currentBlock.style?.spacingAfter === undefined ? 16 : 0;
      this.appendBlock(
        {
          ...currentBlock,
          id: `${currentBlock.id}-part-${partIndex}`,
          sourceBlockId: currentBlock.sourceBlockId ?? currentBlock.id,
          items: currentBlock.items.slice(startItemIndex, itemIndex),
          continuesOnNext: !isFinalPart,
          estimatedHeight: undefined,
        },
        itemsHeight + partOuterSpacing,
      );
      if (!isFinalPart) this.flushPage();
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
      const tableTopOffset = currentBlock.width
        ? DOC_STRUCTURED_TABLE_TOP_OFFSET
        : 0;
      availableHeight = Math.max(
        0,
        availableHeight - partSpacingBefore - tableTopOffset,
      );
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
          false,
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
        partSpacingBefore + tableTopOffset + rowsHeight + partSpacingAfter,
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
  measuredBlockHeights?: ReadonlyMap<string, number>,
): PaginatedDocPage[] {
  return new DocPaginationState(
    page,
    contentWidth,
    measuredBlockHeights,
  ).append(blocks, true);
}
