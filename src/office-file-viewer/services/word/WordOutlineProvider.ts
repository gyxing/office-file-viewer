import type { WordOutlineItem, WordOutlineTreeNode } from './types';

/** 大纲 Provider 的轻量版本快照。 */
export type WordOutlineProviderSnapshot = {
  /** 数据源变更时递增的修订号。 */
  revision: number;
  /** 当前集合或范围包含的项目数量。 */
  count: number;
  /** 当前任务是否已经完整处理。 */
  complete: boolean;
};

/** 为普通和渐进 Word 文档提供统一的按层大纲访问能力。 */
export interface WordOutlineProvider {
  /** 返回当前可观察状态的只读快照。 */
  getSnapshot(): WordOutlineProviderSnapshot;
  /** 订阅状态快照变化，并返回取消订阅函数。 */
  subscribe(listener: () => void): () => void;
  /** 返回大纲中的根节点。 */
  getRoots(maxDepth?: number): WordOutlineTreeNode[];
  /** 返回指定大纲节点的直接子节点。 */
  getChildren(nodeId: string): Promise<WordOutlineTreeNode[]>;
  /** 返回指定大纲节点从根到父节点的祖先链。 */
  getAncestors(nodeId: string): Promise<string[]>;
  /** 定位指定大纲节点对应的文档位置。 */
  locate(nodeId: string): WordOutlineItem | undefined;
}

/** 渐进解析器向大纲 Provider 追加批次时使用的写入接口。 */
export interface ProgressiveWordOutlineProvider extends WordOutlineProvider {
  /** 向渐进式大纲追加一批节点。 */
  append(items: readonly WordOutlineItem[]): void;
  /** 通知接收方增量输出已经结束。 */
  complete(): void;
}
