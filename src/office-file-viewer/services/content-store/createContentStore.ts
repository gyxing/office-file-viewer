import { IndexedDbContentStore } from './IndexedDbContentStore';
import { MemoryContentStore } from './MemoryContentStore';
import type { OfficeContentRecord, OfficeContentStore } from './types';

export type CreateContentStoreOptions<TValue> = {
  sessionId: string;
  namespace: string;
  maxMemoryBytes: number;
  estimateSize(value: TValue): number;
  /** IndexedDB 降级时接收一次诊断信息，不影响继续预览。 */
  onWarning?: (error: unknown) => void;
};

/** 创建以内存 LRU 为热层、IndexedDB 为冷层的会话内容 Store。 */
export function createContentStore<TMeta, TValue>(
  options: CreateContentStoreOptions<TValue>,
): OfficeContentStore<TMeta, TValue> {
  const memory = new MemoryContentStore<TMeta, TValue>({
    maxBytes: options.maxMemoryBytes,
    estimateSize: options.estimateSize,
  });
  let cold: IndexedDbContentStore<TMeta, TValue> | undefined =
    typeof indexedDB === 'undefined'
      ? undefined
      : new IndexedDbContentStore<TMeta, TValue>(
          options.sessionId,
          options.namespace,
        );
  let warned = false;
  let disposePromise: Promise<void> | undefined;

  const disableCold = (error: unknown) => {
    const failedCold = cold;
    cold = undefined;
    if (!warned) {
      warned = true;
      options.onWarning?.(error);
    }
    void failedCold?.dispose();
  };

  return {
    getMeta: (key) => memory.getMeta(key),
    async get(key, signal) {
      const hot = await memory.get(key, signal);
      if (hot?.value !== undefined || !cold) return hot;
      try {
        const stored = await cold.get(key, signal);
        if (!stored) return hot;
        await memory.hydrate(stored);
        return stored;
      } catch (error) {
        if (error instanceof Error && error.name === 'AbortError') throw error;
        disableCold(error);
        return hot;
      }
    },
    async put(record: OfficeContentRecord<TMeta, TValue>) {
      await memory.put(record);
      if (!cold) return;
      try {
        await cold.put(record);
      } catch (error) {
        disableCold(error);
      }
    },
    pin: (keys) => memory.pin(keys),
    async delete(key) {
      await memory.delete(key);
      if (!cold) return;
      try {
        await cold.delete(key);
      } catch (error) {
        disableCold(error);
      }
    },
    dispose() {
      if (disposePromise) return disposePromise;
      const activeCold = cold;
      cold = undefined;
      disposePromise = (async () => {
        await memory.dispose();
        await activeCold?.dispose();
      })();
      return disposePromise;
    },
  };
}
