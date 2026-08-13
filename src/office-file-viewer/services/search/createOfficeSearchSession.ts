import { isOfficeSearchAbortError } from './OfficeSearchProvider';
import type {
  OfficeSearchBatchEmitter,
  OfficeSearchProvider,
  OfficeSearchQuery,
  OfficeSearchSession,
} from './types';

/** 为单个文档创建查询版本隔离且可取消的搜索会话。 */
export function createOfficeSearchSession(
  provider: OfficeSearchProvider,
): OfficeSearchSession {
  let nextQueryId = 1;
  let activeToken = 0;
  let activeController: AbortController | undefined;
  let disposed = false;

  const cancel = () => {
    activeToken += 1;
    activeController?.abort();
    activeController = undefined;
  };

  return {
    async search(query: OfficeSearchQuery, emit: OfficeSearchBatchEmitter) {
      if (disposed) throw new Error('Office 搜索会话已经释放');
      cancel();
      const queryId = nextQueryId;
      nextQueryId += 1;
      const token = activeToken;
      const controller = new AbortController();
      activeController = controller;

      try {
        await provider.search(
          query,
          (progress) => {
            if (
              disposed ||
              controller.signal.aborted ||
              token !== activeToken
            ) {
              return;
            }
            emit({ ...progress, queryId });
          },
          controller.signal,
        );
      } catch (error) {
        if (!isOfficeSearchAbortError(error)) throw error;
      } finally {
        if (activeController === controller) activeController = undefined;
      }
      return queryId;
    },
    cancel,
    dispose() {
      if (disposed) return;
      disposed = true;
      cancel();
    },
  };
}
