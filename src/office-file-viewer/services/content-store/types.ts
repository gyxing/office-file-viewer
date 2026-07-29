/** 描述 Store 中带版本的内容、元数据和更新时间。 */
export type OfficeContentRecord<TMeta, TValue> = {
  key: string;
  revision: number;
  meta: TMeta;
  value?: TValue;
  updatedAt: number;
};

/** 表示不常驻具体内容值的小型元数据记录。 */
export type OfficeContentMetaRecord<TMeta> = Omit<
  OfficeContentRecord<TMeta, unknown>,
  'value'
>;

/** 提供版本保护、冷热读取、固定和幂等释放能力。 */
export interface OfficeContentStore<TMeta, TValue> {
  getMeta(key: string): OfficeContentMetaRecord<TMeta> | undefined;
  get(
    key: string,
    signal?: AbortSignal,
  ): Promise<OfficeContentRecord<TMeta, TValue> | undefined>;
  put(record: OfficeContentRecord<TMeta, TValue>): Promise<void>;
  pin(keys: readonly string[]): () => void;
  delete(key: string): Promise<void>;
  dispose(): Promise<void>;
}

/** 内容版本或 Store 生命周期错误使用的稳定结构。 */
export type OfficeContentStoreError = Error & {
  code: 'STALE_CONTENT_REVISION' | 'CONTENT_STORE_DISPOSED';
};

export type MemoryContentStoreOptions<TValue> = {
  maxBytes: number;
  estimateSize(value: TValue): number;
};
