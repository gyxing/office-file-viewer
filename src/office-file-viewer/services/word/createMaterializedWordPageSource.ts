import type {
  WordPageMeta,
  WordPageSource,
  WordPageSourceSnapshot,
} from './WordPageSource';

export type MaterializedWordPageSourceOptions<TPage> = {
  getId(page: TPage, index: number): string;
  getEstimatedContentHeight(page: TPage, index: number): number;
  getSourceBlockIds(page: TPage, index: number): readonly string[];
};

function throwIfAborted(signal?: AbortSignal) {
  if (!signal?.aborted) return;
  throw new DOMException('操作已取消', 'AbortError');
}

/** 将普通页面数组适配为 Source；不创建 IndexedDB、LRU 或页面副本。 */
export function createMaterializedWordPageSource<TPage>(
  pages: readonly TPage[],
  options: MaterializedWordPageSourceOptions<TPage>,
): WordPageSource<TPage> {
  const metas: WordPageMeta[] = pages.map((page, index) => ({
    id: options.getId(page, index),
    index,
    revision: 1,
    status: 'ready',
    estimatedContentHeight: options.getEstimatedContentHeight(page, index),
    sourceBlockIds: options.getSourceBlockIds(page, index),
  }));
  const blockPageIndex = new Map<string, number>();
  metas.forEach((meta) =>
    meta.sourceBlockIds.forEach((id) => blockPageIndex.set(id, meta.index)),
  );
  const snapshot: WordPageSourceSnapshot = {
    revision: 1,
    pageCount: pages.length,
    pages: metas,
  };

  return {
    getSnapshot: () => snapshot,
    subscribe: () => () => undefined,
    getPage: async (index, signal) => {
      throwIfAborted(signal);
      return pages[index];
    },
    ensureRange: async (_start, _end, signal) => {
      throwIfAborted(signal);
    },
    prioritizeBlock: async (blockId, signal) => {
      throwIfAborted(signal);
      return blockPageIndex.get(blockId) ?? -1;
    },
    retainRange: () => () => undefined,
    retry: () => undefined,
    dispose: () => Promise.resolve(),
  };
}
