/** 带修订号和轻量元数据的 Office 内容缓存记录。 */
export type OfficeContentRecord<TMeta, TValue> = {
  /** 用于稳定识别或缓存当前项目的键。 */
  key: string;
  /** 数据源变更时递增的修订号。 */
  revision: number;
  /** 无需加载完整内容即可读取的轻量元数据。 */
  meta: TMeta;
  /** 按键和修订号缓存的实际内容。 */
  value?: TValue;
  /** 内容记录最近一次更新的时间戳。 */
  updatedAt: number;
};

/** 表示不常驻具体内容值的小型元数据记录。 */
export type OfficeContentMetaRecord<TMeta> = Omit<
  OfficeContentRecord<TMeta, unknown>,
  'value'
>;

/** 提供版本保护、冷热读取、固定和幂等释放能力。 */
export interface OfficeContentStore<TMeta, TValue> {
  /** 读取指定缓存项的元数据。 */
  getMeta(key: string): OfficeContentMetaRecord<TMeta> | undefined;
  /** 读取指定键对应的缓存内容。 */
  get(
    key: string,
    signal?: AbortSignal,
  ): Promise<OfficeContentRecord<TMeta, TValue> | undefined>;
  /** 写入或覆盖指定键对应的缓存内容。 */
  put(record: OfficeContentRecord<TMeta, TValue>): Promise<void>;
  /** 更新缓存项的固定状态，避免其被容量淘汰。 */
  pin(keys: readonly string[]): () => void;
  /** 删除指定键对应的缓存内容。 */
  delete(key: string): Promise<void>;
  /** 幂等释放当前对象持有的资源和订阅。 */
  dispose(): Promise<void>;
}

/** 内容版本或 Store 生命周期错误使用的稳定结构。 */
export type OfficeContentStoreError = Error & {
  /** 供程序识别当前情况的稳定代码。 */
  code: 'STALE_CONTENT_REVISION' | 'CONTENT_STORE_DISPOSED';
};

/** 内存内容存储的容量选项。 */
export type MemoryContentStoreOptions<TValue> = {
  /** 当前存储允许占用的最大字节数。 */
  maxBytes: number;
  /** 估算单个缓存值占用的字节数。 */
  estimateSize(value: TValue): number;
};
