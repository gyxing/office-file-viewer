import type { WordRevisionMode } from '../annotations/types';
import type {
  DocxBlock,
  DocxInline,
  DocxPageContent,
  DocxParagraphBlock,
  DocxShape,
  DocxShapeItem,
  DocxTableBlock,
} from './types';

/** 判断修订内容在指定 Word 投影模式下是否可见。 */
function isInlineVisible(inline: DocxInline, mode: WordRevisionMode) {
  const revisions = inline.review?.revisions ?? [];
  if (mode === 'markup') return true;
  if (mode === 'final') {
    return !revisions.some(
      (revision) => revision.kind === 'delete' || revision.kind === 'move-from',
    );
  }
  return !revisions.some(
    (revision) => revision.kind === 'insert' || revision.kind === 'move-to',
  );
}

/** 在 original 模式下恢复格式修订前的文字样式。 */
function projectInline(
  inline: DocxInline,
  mode: WordRevisionMode,
): DocxInline | undefined {
  if (!isInlineVisible(inline, mode)) return undefined;
  let projected = inline;
  if (
    mode === 'original' &&
    inline.type === 'text' &&
    inline.review?.originalStyle
  ) {
    projected = { ...inline, style: inline.review.originalStyle };
  }
  if (projected.type !== 'shape') return projected;
  return {
    ...projected,
    shape: projectShape(projected.shape, mode),
  };
}

/** 计算投影后段落参与目录、搜索和分页的可见文本。 */
function textFromProjectedInlines(inlines: readonly DocxInline[]) {
  return inlines
    .map((inline) =>
      inline.type === 'text' ? inline.text : inline.type === 'tab' ? '\t' : '',
    )
    .join('');
}

/** 投影段落内的修订内容，并同步刷新段落可见文本。 */
function projectParagraph(
  block: DocxParagraphBlock,
  mode: WordRevisionMode,
): DocxParagraphBlock {
  const inlines = block.inlines.flatMap((inline) => {
    const projected = projectInline(inline, mode);
    return projected ? [projected] : [];
  });
  const text = textFromProjectedInlines(inlines);
  const deletedParagraphMark = block.review?.revisions?.some(
    (revision) => revision.kind === 'delete' || revision.kind === 'move-from',
  );
  return {
    ...block,
    inlines,
    text,
    revisionHidden: mode === 'final' && deletedParagraphMark && !text,
    style:
      mode === 'original' && block.review?.originalStyle
        ? block.review.originalStyle
        : block.style,
  };
}

/** 递归投影表格单元格内的段落、嵌套表格和形状。 */
function projectTable(
  block: DocxTableBlock,
  mode: WordRevisionMode,
): DocxTableBlock {
  return {
    ...block,
    rows: block.rows.map((row) => ({
      ...row,
      cells: row.cells.map((cell) => ({
        ...cell,
        blocks: cell.blocks.map((child) => projectDocxBlock(child, mode)),
      })),
    })),
  };
}

/** 投影形状内部的文本块，保持几何和绘图资源引用不变。 */
function projectShapeItem(
  item: DocxShapeItem,
  mode: WordRevisionMode,
): DocxShapeItem {
  return {
    ...item,
    blocks: item.blocks?.map((block) => projectDocxBlock(block, mode)),
    paragraphs: item.paragraphs?.map((paragraph) =>
      projectParagraph(paragraph, mode),
    ),
  };
}

/** 投影复合形状内部的全部文字内容。 */
function projectShape(shape: DocxShape, mode: WordRevisionMode): DocxShape {
  return {
    ...shape,
    items: shape.items.map((item) => projectShapeItem(item, mode)),
  };
}

/** 将单个 DOCX 块投影为指定修订视图。 */
export function projectDocxBlock(
  block: DocxBlock,
  mode: WordRevisionMode,
): DocxBlock {
  if (block.type === 'paragraph') return projectParagraph(block, mode);
  if (block.type === 'table') return projectTable(block, mode);
  return block;
}

/** 将页面正文和页眉中的修订内容投影为指定视图。 */
export function projectDocxPageContent(
  page: DocxPageContent,
  mode: WordRevisionMode,
): DocxPageContent {
  const projectRegion = (
    region: DocxPageContent['headers'],
  ): DocxPageContent['headers'] =>
    region
      ? {
          default: region.default?.map((block) =>
            projectDocxBlock(block, mode),
          ),
          first: region.first?.map((block) => projectDocxBlock(block, mode)),
          even: region.even?.map((block) => projectDocxBlock(block, mode)),
        }
      : undefined;
  return {
    ...page,
    blocks: page.blocks.map((block) => projectDocxBlock(block, mode)),
    headers: projectRegion(page.headers),
  };
}
