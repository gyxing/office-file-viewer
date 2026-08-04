import { createContentAbortError, createContentStoreError } from './errors';
import type {
  MemoryContentStoreOptions,
  OfficeContentMetaRecord,
  OfficeContentRecord,
  OfficeContentStore,
} from './types';

/** 内存内容存储中的资源条目。 */
type MemoryEntry<TMeta, TValue> = {
  /** 当前内存条目保存的内容记录。 */
  record: OfficeContentRecord<TMeta, TValue>;
  /** 当前数据占用的空间大小。 */
  size: number;
};

/** 以内存字节预算管理热内容，并保留轻量 revision 元数据。 */
export class MemoryContentStore<TMeta, TValue>
  implements OfficeContentStore<TMeta, TValue>
{
  private readonly entries = new Map<string, MemoryEntry<TMeta, TValue>>();
  /** 仅记录仍驻留内存的内容，并利用 Map 插入顺序维护 LRU。 */
  private readonly residentKeys = new Map<string, undefined>();
  private readonly pinCounts = new Map<string, number>();
  private usedBytes = 0;
  private disposed = false;
  private disposePromise: Promise<void> | undefined;

  constructor(private readonly options: MemoryContentStoreOptions<TValue>) {
    if (!Number.isFinite(options.maxBytes) || options.maxBytes < 0) {
      throw new RangeError('ContentStore 内存预算必须是非负有限数');
    }
  }

  private ensureAvailable(signal?: AbortSignal) {
    if (this.disposed) {
      throw createContentStoreError(
        'CONTENT_STORE_DISPOSED',
        'ContentStore 已释放',
      );
    }
    if (signal?.aborted) throw createContentAbortError();
  }

  private evictIfNeeded() {
    while (this.usedBytes > this.options.maxBytes) {
      let candidateKey: string | undefined;
      for (const key of this.residentKeys.keys()) {
        if ((this.pinCounts.get(key) ?? 0) > 0) continue;
        candidateKey = key;
        break;
      }
      if (candidateKey === undefined) return;

      this.residentKeys.delete(candidateKey);
      const entry = this.entries.get(candidateKey);
      if (!entry || entry.record.value === undefined) continue;
      this.usedBytes -= entry.size;
      this.entries.set(candidateKey, {
        record: {
          key: entry.record.key,
          revision: entry.record.revision,
          meta: entry.record.meta,
          updatedAt: entry.record.updatedAt,
        },
        size: 0,
      });
    }
  }

  /** 把最近访问的驻留内容移动到 LRU 队尾。 */
  private markResidentUsed(key: string) {
    this.residentKeys.delete(key);
    this.residentKeys.set(key, undefined);
  }

  getMeta(key: string): OfficeContentMetaRecord<TMeta> | undefined {
    this.ensureAvailable();
    const record = this.entries.get(key)?.record;
    if (!record) return undefined;
    return {
      key: record.key,
      revision: record.revision,
      meta: record.meta,
      updatedAt: record.updatedAt,
    };
  }

  async get(key: string, signal?: AbortSignal) {
    this.ensureAvailable(signal);
    const entry = this.entries.get(key);
    if (!entry) return undefined;
    if (entry.record.value !== undefined) {
      this.markResidentUsed(key);
    }
    return { ...entry.record };
  }

  async put(record: OfficeContentRecord<TMeta, TValue>) {
    this.ensureAvailable();
    const current = this.entries.get(record.key);
    if (current && record.revision <= current.record.revision) {
      throw createContentStoreError(
        'STALE_CONTENT_REVISION',
        `内容 ${record.key} 的 revision 必须递增`,
      );
    }

    const size =
      record.value === undefined
        ? 0
        : Math.max(0, this.options.estimateSize(record.value));
    if (!Number.isFinite(size)) {
      throw new RangeError('ContentStore 内容大小估算必须是有限数');
    }
    if (current) {
      this.usedBytes -= current.size;
      this.residentKeys.delete(record.key);
    }
    this.usedBytes += size;
    this.entries.set(record.key, {
      record: { ...record },
      size,
    });
    if (record.value !== undefined) this.markResidentUsed(record.key);
    this.evictIfNeeded();
  }

  /** 把冷层读取的相同 revision 内容恢复到热层，不放宽正常写入规则。 */
  async hydrate(record: OfficeContentRecord<TMeta, TValue>) {
    this.ensureAvailable();
    const current = this.entries.get(record.key);
    if (!current || current.record.revision < record.revision) {
      await this.put(record);
      return;
    }
    if (
      current.record.revision > record.revision ||
      current.record.value !== undefined
    ) {
      return;
    }

    const size =
      record.value === undefined
        ? 0
        : Math.max(0, this.options.estimateSize(record.value));
    if (!Number.isFinite(size)) {
      throw new RangeError('ContentStore 内容大小估算必须是有限数');
    }
    current.record = { ...record };
    current.size = size;
    this.usedBytes += size;
    if (record.value !== undefined) this.markResidentUsed(record.key);
    this.evictIfNeeded();
  }

  pin(keys: readonly string[]) {
    this.ensureAvailable();
    const uniqueKeys = [...new Set(keys)];
    uniqueKeys.forEach((key) => {
      this.pinCounts.set(key, (this.pinCounts.get(key) ?? 0) + 1);
    });
    let released = false;
    return () => {
      if (released) return;
      released = true;
      uniqueKeys.forEach((key) => {
        const next = (this.pinCounts.get(key) ?? 1) - 1;
        if (next > 0) this.pinCounts.set(key, next);
        else this.pinCounts.delete(key);
      });
      this.evictIfNeeded();
    };
  }

  async delete(key: string) {
    this.ensureAvailable();
    const entry = this.entries.get(key);
    if (entry) this.usedBytes -= entry.size;
    this.entries.delete(key);
    this.residentKeys.delete(key);
    this.pinCounts.delete(key);
  }

  dispose() {
    if (this.disposePromise) return this.disposePromise;
    this.disposed = true;
    this.entries.clear();
    this.residentKeys.clear();
    this.pinCounts.clear();
    this.usedBytes = 0;
    this.disposePromise = Promise.resolve();
    return this.disposePromise;
  }
}
