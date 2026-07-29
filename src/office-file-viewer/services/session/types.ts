/** 描述由文档会话统一持有和释放的资源。 */
export type OfficeSessionResource = {
  /** 释放资源；实现必须允许会话只调用一次。 */
  dispose(): void | Promise<void>;
};

/** 管理单次文档加载的取消信号、稳定标识和资源所有权。 */
export type OfficeDocumentSession = {
  /** 与文件名和模型引用无关的稳定会话标识。 */
  readonly id: string;
  /** 贯穿解析、资源读取和渲染任务的统一取消信号。 */
  readonly signal: AbortSignal;
  /** 注册会话资源，返回的函数仅取消注册，不主动释放资源。 */
  register(resource: OfficeSessionResource): () => void;
  /** 把会话资源的释放责任转移给最终文档或工作簿。 */
  transferTo(owner: object): void;
  /** 取消当前会话仍在运行的任务。 */
  abort(reason?: unknown): void;
  /** 幂等释放当前会话持有的全部资源。 */
  dispose(): Promise<void>;
};
