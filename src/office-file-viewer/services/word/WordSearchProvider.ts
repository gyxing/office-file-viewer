import type { DocBlock } from '../doc/types';
import type { DocxBlock, DocxParagraphBlock } from '../docx/types';
import {
  createOfficeSearchAbortError,
  OfficeSearchBatchWriter,
  throwIfOfficeSearchAborted,
} from '../search/OfficeSearchProvider';
import {
  createSearchPreviewText,
  findSearchMatches,
  normalizeSearchText,
} from '../search/normalizeSearchText';
import type {
  OfficeSearchProgressEmitter,
  OfficeSearchProvider,
  OfficeSearchQuery,
  OfficeSearchResult,
} from '../search/types';

/** Word 搜索扫描使用的最小正文块快照。 */
export type WordSearchBlock = Readonly<{
  /** 正文块的稳定标识。 */
  blockId: string;
  /** 当前块参与搜索的可见文本。 */
  text: string;
  /** 已知时记录零基页面索引。 */
  pageIndex?: number;
}>;

/** Word 搜索提供器的页面定位选项。 */
export type WordSearchProviderOptions = Readonly<{
  /** 渐进分页完成后按块标识补充实际页面索引。 */
  resolvePageIndex?(blockId: string): number | undefined;
}>;

function docBlockText(block: DocBlock) {
  if (block.type === 'paragraph') return block.text;
  if (block.type === 'list')
    return block.items.map((item) => item.text).join('\n');
  return block.rows
    .map((row) => row.cells.map((cell) => cell.text).join('\t'))
    .join('\n');
}

function docxParagraphShapeText(block: DocxParagraphBlock) {
  return block.inlines
    .flatMap((inline) => (inline.type === 'shape' ? inline.shape.items : []))
    .map((item) => item.blocks ?? item.paragraphs ?? [])
    .map((blocks) => blocks.map(docxBlockText).filter(Boolean).join('\n'))
    .filter(Boolean);
}

function docxBlockText(block: DocxBlock): string {
  if (block.type === 'chart') return '';
  if (block.type === 'paragraph') {
    return [block.text, ...docxParagraphShapeText(block)]
      .filter(Boolean)
      .join('\n');
  }
  return block.rows
    .map((row) =>
      row.cells
        .map((cell) =>
          cell.blocks.map(docxBlockText).filter(Boolean).join('\n'),
        )
        .join('\t'),
    )
    .join('\n');
}

/** 将 DOC 标准模型转换为不复制样式和资源的搜索块。 */
export function collectDocSearchBlocks(
  blocks: readonly DocBlock[],
): WordSearchBlock[] {
  return blocks.map((block) => ({
    blockId: block.sourceBlockId ?? block.id,
    text: docBlockText(block),
  }));
}

/** 将 DOCX 标准模型转换为不复制样式和资源的搜索块。 */
export function collectDocxSearchBlocks(
  blocks: readonly DocxBlock[],
): WordSearchBlock[] {
  return blocks.map((block) => ({
    blockId: block.sourceBlockId ?? block.id,
    text: docxBlockText(block),
  }));
}

/** 支持解析期间持续追加正文块的 Word 增量搜索提供器。 */
export class WordSearchProvider implements OfficeSearchProvider {
  readonly kind = 'word' as const;
  private readonly blocks: WordSearchBlock[] = [];
  private readonly blockIndexById = new Map<string, number>();
  private readonly waiters = new Set<() => void>();
  private revision = 0;
  private completed = false;
  private disposed = false;

  constructor(private readonly options: WordSearchProviderOptions = {}) {}

  /** 追加解析器新产生的正文块；重复块只更新其页面定位。 */
  append(blocks: readonly WordSearchBlock[]) {
    if (this.completed || this.disposed) return;
    blocks.forEach((block) => {
      const existingIndex = this.blockIndexById.get(block.blockId);
      if (existingIndex !== undefined) {
        const existing = this.blocks[existingIndex];
        this.blocks[existingIndex] = {
          ...existing,
          pageIndex: block.pageIndex ?? existing.pageIndex,
        };
        return;
      }
      this.blockIndexById.set(block.blockId, this.blocks.length);
      this.blocks.push(block);
    });
    this.notifyChange();
  }

  /** 在分页结果可用后补充块所在页面，不重新扫描正文内容。 */
  setPageIndex(blockId: string, pageIndex: number) {
    const index = this.blockIndexById.get(blockId);
    if (index === undefined) return;
    this.blocks[index] = { ...this.blocks[index], pageIndex };
  }

  /** 标记正文解析完成并唤醒等待末批数据的查询。 */
  complete() {
    if (this.completed) return;
    this.completed = true;
    this.notifyChange();
  }

  async search(
    query: OfficeSearchQuery,
    emit: OfficeSearchProgressEmitter,
    signal: AbortSignal,
  ) {
    throwIfOfficeSearchAborted(signal);
    const writer = new OfficeSearchBatchWriter(
      emit,
      signal,
      this.blocks.length,
    );
    if (!normalizeSearchText(query.text, query.matchCase).text) {
      writer.complete();
      return;
    }

    let cursor = 0;
    while (true) {
      throwIfOfficeSearchAborted(signal);
      const observedRevision = this.revision;
      const availableLength = this.blocks.length;
      writer.setTotal(availableLength);

      while (cursor < availableLength) {
        const block = this.blocks[cursor];
        const pageIndex =
          block.pageIndex ?? this.options.resolvePageIndex?.(block.blockId);
        const items: OfficeSearchResult[] = findSearchMatches(
          block.text,
          query,
        ).map(({ startOffset, endOffset }) => ({
          id: `word:${block.blockId}:${startOffset}:${endOffset}`,
          matchText: block.text.slice(startOffset, endOffset),
          previewText: createSearchPreviewText(
            block.text,
            startOffset,
            endOffset,
          ),
          target: {
            kind: 'word',
            blockId: block.blockId,
            ...(pageIndex === undefined || pageIndex < 0 ? {} : { pageIndex }),
            startOffset,
            endOffset,
          },
        }));
        cursor += 1;
        await writer.append(items);
      }

      if ((this.completed || this.disposed) && cursor >= this.blocks.length) {
        break;
      }
      await this.waitForChange(observedRevision, signal);
    }
    writer.setTotal(this.blocks.length);
    writer.complete();
  }

  /** 释放渐进等待器；已复制到查询中的块会自然完成当前批次。 */
  dispose() {
    if (this.disposed) return;
    this.disposed = true;
    this.completed = true;
    this.notifyChange();
  }

  private notifyChange() {
    this.revision += 1;
    const waiters = [...this.waiters];
    this.waiters.clear();
    waiters.forEach((resolve) => resolve());
  }

  private waitForChange(revision: number, signal: AbortSignal) {
    if (this.revision !== revision || this.completed || this.disposed) {
      return Promise.resolve();
    }
    return new Promise<void>((resolve, reject) => {
      let finish: () => void;
      const abort = () => {
        this.waiters.delete(finish);
        signal.removeEventListener('abort', abort);
        reject(createOfficeSearchAbortError());
      };
      finish = () => {
        signal.removeEventListener('abort', abort);
        this.waiters.delete(finish);
        resolve();
      };
      this.waiters.add(finish);
      signal.addEventListener('abort', abort, { once: true });
    });
  }
}

/** 为已完整解析的 Word 正文创建立即完成的内存搜索提供器。 */
export function createMaterializedWordSearchProvider(
  blocks: readonly WordSearchBlock[],
  options: WordSearchProviderOptions = {},
) {
  const provider = new WordSearchProvider(options);
  provider.append(blocks);
  provider.complete();
  return provider;
}
