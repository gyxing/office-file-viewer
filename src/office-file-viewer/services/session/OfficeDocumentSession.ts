import { attachDocumentSession } from './documentResourceRegistry';
import type { OfficeDocumentSession, OfficeSessionResource } from './types';

let fallbackSessionSequence = 0;

/** 创建不依赖文件名的会话标识，兼容缺少 randomUUID 的浏览器。 */
function createDocumentSessionId() {
  if (
    typeof crypto !== 'undefined' &&
    typeof crypto.randomUUID === 'function'
  ) {
    return crypto.randomUUID();
  }
  fallbackSessionSequence += 1;
  return `office-session-${Date.now()}-${fallbackSessionSequence}`;
}

/** 创建单次文档加载使用的稳定资源会话。 */
export function createOfficeDocumentSession(): OfficeDocumentSession {
  const controller = new AbortController();
  const resources = new Set<OfficeSessionResource>();
  let disposePromise: Promise<void> | undefined;

  const session: OfficeDocumentSession = {
    id: createDocumentSessionId(),
    signal: controller.signal,
    register(resource) {
      if (disposePromise) {
        void resource.dispose();
        return () => undefined;
      }
      resources.add(resource);
      return () => resources.delete(resource);
    },
    transferTo(owner) {
      if (disposePromise) {
        throw new Error('已释放的文档会话不能转移所有权');
      }
      attachDocumentSession(owner, session);
    },
    abort(reason) {
      if (!controller.signal.aborted) controller.abort(reason);
    },
    dispose() {
      if (disposePromise) return disposePromise;
      controller.abort();
      const ownedResources = [...resources].reverse();
      resources.clear();
      disposePromise = (async () => {
        let firstError: unknown;
        for (const resource of ownedResources) {
          try {
            await resource.dispose();
          } catch (error) {
            firstError ??= error;
          }
        }
        if (firstError) throw firstError;
      })();
      return disposePromise;
    },
  };

  return session;
}
