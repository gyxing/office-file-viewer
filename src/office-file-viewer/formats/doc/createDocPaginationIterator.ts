import type { DocBlock, DocPage } from '../../services/doc/types';
import {
  createTimeBudget,
  throwIfAborted,
  yieldToMainThread,
} from '../../services/performance/mainThreadScheduler';
import { DocPaginationState, type PaginatedDocPage } from '../../services/doc/docPagination';

/** 按页面批次组织的 DOC 内容块集合。 */
type DocBlockBatches = AsyncIterable<DocBlock[]> | Iterable<DocBlock[]>;

/** 分批消费 DOC blocks，并在每个完成页和 8ms 时间片边界让出主线程。 */
export async function* createDocPaginationIterator(
  blocks: DocBlockBatches,
  page: DocPage,
  contentWidth: number,
  signal?: AbortSignal,
): AsyncGenerator<PaginatedDocPage> {
  const state = new DocPaginationState(page, contentWidth);
  const budget = createTimeBudget();
  for await (const batch of blocks) {
    for (const block of batch) {
      throwIfAborted(signal);
      const pages = state.append([block]);
      for (const completedPage of pages) yield completedPage;
      if (budget.shouldYield()) {
        await yieldToMainThread(signal);
        budget.reset();
      }
    }
  }
  for (const completedPage of state.append([], true)) {
    throwIfAborted(signal);
    yield completedPage;
  }
}
