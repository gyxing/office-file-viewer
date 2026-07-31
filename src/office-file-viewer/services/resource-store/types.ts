/** 描述可直接使用或按需物化的 Office 资源来源。 */
export type OfficeResourceSource =
  | {
      /** 当前模型对应的 Office 内容类型。 */
      kind: 'url';
      /** 资源访问地址。 */
      url: string;
    }
  | {
      /** 当前模型对应的 Office 内容类型。 */
      kind: 'lazy';
      /** 在所属集合中的唯一标识。 */
      id: string;
      /** 资源负载的 MIME 类型。 */
      mimeType: string;
      /** 当前数据占用的空间大小。 */
      size: number;
      /** 按需加载资源并返回可读取的 Blob。 */
      load(signal?: AbortSignal): Promise<Blob>;
    };

/** 管理资源并发去重、Object URL 引用和统一释放。 */
export interface OfficeResourceStore {
  /** 获取资源并增加其引用计数。 */
  acquire(source: OfficeResourceSource, signal?: AbortSignal): Promise<string>;
  /** 减少资源引用计数，并在无人使用时释放资源。 */
  release(source: OfficeResourceSource): void;
  /** 幂等释放当前对象持有的资源和订阅。 */
  dispose(): Promise<void>;
}

/** Office 资源存储的内存控制选项。 */
export type OfficeResourceStoreOptions = {
  /** 未引用 Blob 和 Object URL 的内存预算。 */
  maxUnusedBytes?: number;
};
