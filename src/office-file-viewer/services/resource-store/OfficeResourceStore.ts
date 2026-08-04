import type {
  OfficeResourceSource,
  OfficeResourceStore,
  OfficeResourceStoreOptions,
} from './types';

/** 资源存储允许暂存的未引用资源大小，单位为字节。 */
const DEFAULT_UNUSED_RESOURCE_BYTES = 32 * 1024 * 1024;

/** 等待首次读取时再加载的资源源。 */
type LazySource = Extract<
  OfficeResourceSource,
  {
    /** 标识该资源需要等到实际使用时再加载。 */
    kind: 'lazy';
  }
>;

/** 按需资源的引用计数、缓存和加载状态。 */
type ResourceEntry = {
  /** 尚未物化的资源加载定义。 */
  source: LazySource;
  /** 当前资源被预览节点持有的引用数量。 */
  references: number;
  /** 当前数据占用的空间大小。 */
  size: number;
  /** 完成按需读取后缓存的二进制对象。 */
  blob?: Blob;
  /** 资源访问地址。 */
  url?: string;
  /** 用于控制当前视图或任务的控制器。 */
  controller?: AbortController;
  /** 加载状态相关文案。 */
  loading?: Promise<string>;
};

function createAbortError() {
  const error = new Error('资源加载已取消');
  error.name = 'AbortError';
  return error;
}

function waitForResource(
  promise: Promise<string>,
  signal: AbortSignal | undefined,
  onAbort: () => void,
) {
  if (!signal) return promise;
  if (signal.aborted) {
    onAbort();
    return Promise.reject(createAbortError());
  }
  return new Promise<string>((resolve, reject) => {
    const handleAbort = () => {
      onAbort();
      reject(createAbortError());
    };
    signal.addEventListener('abort', handleAbort, { once: true });
    promise.then(
      (url) => {
        signal.removeEventListener('abort', handleAbort);
        resolve(url);
      },
      (error) => {
        signal.removeEventListener('abort', handleAbort);
        reject(error);
      },
    );
  });
}

/** 以稳定资源 ID 管理懒加载 Blob、引用计数和未引用 LRU。 */
export class ManagedOfficeResourceStore implements OfficeResourceStore {
  private readonly entries = new Map<string, ResourceEntry>();
  private readonly inFlight = new Set<Promise<string>>();
  /** 仅保存未引用且已物化的资源，并利用 Map 插入顺序维护 LRU。 */
  private readonly unusedEntries = new Map<string, ResourceEntry>();
  private readonly maxUnusedBytes: number;
  private unusedBytes = 0;
  private disposed = false;
  private disposePromise: Promise<void> | undefined;

  constructor(options: OfficeResourceStoreOptions = {}) {
    this.maxUnusedBytes =
      options.maxUnusedBytes ?? DEFAULT_UNUSED_RESOURCE_BYTES;
    if (!Number.isFinite(this.maxUnusedBytes) || this.maxUnusedBytes < 0) {
      throw new RangeError('资源 Store 预算必须是非负有限数');
    }
  }

  private ensureAvailable(signal?: AbortSignal) {
    if (this.disposed) throw new Error('OfficeResourceStore 已释放');
    if (signal?.aborted) throw createAbortError();
  }

  private revoke(entry: ResourceEntry) {
    if (
      entry.url &&
      typeof URL !== 'undefined' &&
      typeof URL.revokeObjectURL === 'function'
    ) {
      URL.revokeObjectURL(entry.url);
    }
    entry.url = undefined;
    entry.blob = undefined;
    entry.size = 0;
  }

  private evictIfNeeded() {
    while (this.unusedBytes > this.maxUnusedBytes) {
      const candidate = this.unusedEntries.entries().next();
      if (candidate.done) return;
      const [id, entry] = candidate.value;
      this.unusedEntries.delete(id);
      this.unusedBytes -= entry.size;
      if (this.entries.get(id) === entry) this.entries.delete(id);
      this.revoke(entry);
    }
  }

  /** 资源再次被引用时同步移出未使用 LRU。 */
  private removeUnusedEntry(entry: ResourceEntry) {
    if (!this.unusedEntries.delete(entry.source.id)) return;
    this.unusedBytes -= entry.size;
  }

  /** 资源引用归零后加入未使用 LRU，并按预算即时回收。 */
  private addUnusedEntry(entry: ResourceEntry) {
    if (
      entry.references > 0 ||
      entry.loading ||
      !entry.url ||
      this.unusedEntries.has(entry.source.id)
    ) {
      return;
    }
    this.unusedEntries.set(entry.source.id, entry);
    this.unusedBytes += entry.size;
    this.evictIfNeeded();
  }

  private releaseEntry(entry: ResourceEntry) {
    if (entry.references <= 0) return;
    entry.references -= 1;
    if (entry.references > 0) return;
    if (entry.loading) {
      entry.controller?.abort();
      return;
    }
    this.addUnusedEntry(entry);
  }

  private startLoading(entry: ResourceEntry) {
    const controller = new AbortController();
    entry.controller = controller;
    const loading = (async () => {
      const blob = await entry.source.load(controller.signal);
      if (this.disposed || controller.signal.aborted) throw createAbortError();
      if (
        typeof URL === 'undefined' ||
        typeof URL.createObjectURL !== 'function'
      ) {
        throw new Error('当前环境不支持 Blob URL');
      }
      entry.blob = blob;
      entry.size = blob.size;
      entry.url = URL.createObjectURL(blob);
      return entry.url;
    })();
    entry.loading = loading;
    this.inFlight.add(loading);
    void loading
      .catch(() => undefined)
      .finally(() => {
        this.inFlight.delete(loading);
        entry.loading = undefined;
        entry.controller = undefined;
        if (!entry.url) {
          if (this.entries.get(entry.source.id) === entry) {
            this.entries.delete(entry.source.id);
          }
          return;
        }
        if (entry.references === 0) {
          this.addUnusedEntry(entry);
        }
      });
    return loading;
  }

  async acquire(source: OfficeResourceSource, signal?: AbortSignal) {
    this.ensureAvailable(signal);
    if (source.kind === 'url') return source.url;

    let entry = this.entries.get(source.id);
    if (!entry) {
      entry = {
        source,
        references: 0,
        size: 0,
      };
      this.entries.set(source.id, entry);
    } else if (entry.references === 0 && entry.url) {
      this.removeUnusedEntry(entry);
    }
    entry.references += 1;
    if (entry.url) return entry.url;

    const loading = entry.loading ?? this.startLoading(entry);
    try {
      return await waitForResource(loading, signal, () =>
        this.releaseEntry(entry!),
      );
    } catch (error) {
      if (!signal?.aborted) this.releaseEntry(entry);
      throw error;
    }
  }

  release(source: OfficeResourceSource) {
    if (source.kind === 'url' || this.disposed) return;
    const entry = this.entries.get(source.id);
    if (entry) this.releaseEntry(entry);
  }

  dispose() {
    if (this.disposePromise) return this.disposePromise;
    this.disposed = true;
    this.entries.forEach((entry) => entry.controller?.abort());
    const pending = [...this.inFlight];
    this.disposePromise = (async () => {
      await Promise.allSettled(pending);
      this.entries.forEach((entry) => this.revoke(entry));
      this.entries.clear();
      this.inFlight.clear();
      this.unusedEntries.clear();
      this.unusedBytes = 0;
    })();
    return this.disposePromise;
  }
}

/** 创建使用默认 32 MiB 未引用预算的资源 Store。 */
export function createOfficeResourceStore(
  options?: OfficeResourceStoreOptions,
) {
  return new ManagedOfficeResourceStore(options);
}
