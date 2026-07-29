/** 描述可直接使用或按需物化的 Office 资源来源。 */
export type OfficeResourceSource =
  | { kind: 'url'; url: string }
  | {
      kind: 'lazy';
      id: string;
      mimeType: string;
      size: number;
      load(signal?: AbortSignal): Promise<Blob>;
    };

/** 管理资源并发去重、Object URL 引用和统一释放。 */
export interface OfficeResourceStore {
  acquire(source: OfficeResourceSource, signal?: AbortSignal): Promise<string>;
  release(source: OfficeResourceSource): void;
  dispose(): Promise<void>;
}

export type OfficeResourceStoreOptions = {
  /** 未引用 Blob 和 Object URL 的内存预算。 */
  maxUnusedBytes?: number;
};
