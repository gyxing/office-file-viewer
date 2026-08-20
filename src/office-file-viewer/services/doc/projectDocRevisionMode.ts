import type { WordRevisionMode } from '../annotations/types';
import type { DocBlock, DocTextInline, DocTextRunInline } from './types';

/** 判断 DOC/WPS 文字修订在指定投影模式下是否可见。 */
function isTextInlineVisible(inline: DocTextInline, mode: WordRevisionMode) {
  if (inline.type !== 'text' || mode === 'markup') return true;
  const revisions = inline.review?.revisions ?? [];
  if (mode === 'final') {
    return !revisions.some((revision) => revision.kind === 'delete');
  }
  return !revisions.some((revision) => revision.kind === 'insert');
}

/** 投影一组 DOC/WPS 行内节点并刷新可见文本。 */
function projectInlines(
  inlines: readonly DocTextInline[] | undefined,
  mode: WordRevisionMode,
) {
  return (inlines ?? []).filter((inline) => isTextInlineVisible(inline, mode));
}

/** 计算投影后的纯文本内容。 */
function textFromInlines(inlines: readonly DocTextInline[]) {
  return inlines
    .filter((inline): inline is DocTextRunInline => inline.type === 'text')
    .map((inline) => inline.text)
    .join('');
}

/** 将单个 DOC/WPS 块投影为最终态、标记态或原始态。 */
export function projectDocBlockRevisionMode(
  block: DocBlock,
  mode: WordRevisionMode,
): DocBlock {
  if (block.type === 'paragraph') {
    const inlines = projectInlines(block.inlines, mode);
    return { ...block, inlines, text: textFromInlines(inlines) };
  }
  if (block.type === 'list') {
    return {
      ...block,
      items: block.items.map((item) => {
        const inlines = projectInlines(item.inlines, mode);
        return { ...item, inlines, text: textFromInlines(inlines) };
      }),
    };
  }
  return {
    ...block,
    rows: block.rows.map((row) => ({
      ...row,
      cells: row.cells.map((cell) => {
        const inlines = projectInlines(cell.inlines, mode);
        return { ...cell, inlines, text: textFromInlines(inlines) };
      }),
    })),
  };
}

/** 按文档顺序投影全部 DOC/WPS 正文块。 */
export function projectDocBlocksRevisionMode(
  blocks: readonly DocBlock[],
  mode: WordRevisionMode,
) {
  return blocks.map((block) => projectDocBlockRevisionMode(block, mode));
}
