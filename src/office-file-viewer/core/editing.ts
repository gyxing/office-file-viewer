import type {
  OfficeChangeSet,
  OfficeDocumentRuntime,
  OfficeNodeId,
} from './types';

/**
 * 编辑器运行时可执行的一次命令结果。
 * @experimental 未来 Editor 边界契约。
 */
export type OfficeCommand = Readonly<{
  id: string;
  label: string;
  changes: OfficeChangeSet;
}>;

/**
 * 编辑器命令定义，不进入快照或 Worker 消息。
 * @experimental 未来 Editor 边界契约。
 */
export type OfficeCommandDefinition = Readonly<{
  id: string;
  label: string;
  execute(): OfficeCommand;
}>;

/**
 * 可撤销的一次编辑事务。
 * @experimental 未来 Editor 边界契约。
 */
export type OfficeTransaction = Readonly<{
  id: string;
  label: string;
  changes: OfficeChangeSet;
}>;

/**
 * 文档文本或节点的当前选区。
 * @experimental 未来 Editor 边界契约。
 */
export type OfficeSelection = Readonly<{
  nodeId: OfficeNodeId;
  start: number;
  end: number;
}>;

/**
 * 编辑器拥有的撤销/重做运行时边界。
 * @experimental 未来 Editor 边界契约。
 */
export type OfficeHistory = {
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  undo(): void;
  redo(): void;
  push(transaction: OfficeTransaction): void;
  clear(): void;
};

/**
 * 未来 Editor 将格式运行时适配为自己的编辑模型。
 * @experimental 未来 Editor 边界契约。
 */
export type OfficeEditorPluginAdapter<TModel = unknown, TSource = unknown> = {
  pluginId: string;
  createEditorModel(document: OfficeDocumentRuntime<TModel, TSource>): unknown;
  commands: readonly OfficeCommandDefinition[];
};
