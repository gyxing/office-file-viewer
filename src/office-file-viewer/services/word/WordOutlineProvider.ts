import type { WordOutlineItem, WordOutlineTreeNode } from './types';

/** 大纲 Provider 的轻量版本快照。 */
export type WordOutlineProviderSnapshot = {
  revision: number;
  count: number;
  complete: boolean;
};

/** 为普通和渐进 Word 文档提供统一的按层大纲访问能力。 */
export interface WordOutlineProvider {
  getSnapshot(): WordOutlineProviderSnapshot;
  subscribe(listener: () => void): () => void;
  getRoots(maxDepth?: number): WordOutlineTreeNode[];
  getChildren(nodeId: string): Promise<WordOutlineTreeNode[]>;
  getAncestors(nodeId: string): Promise<string[]>;
  locate(nodeId: string): WordOutlineItem | undefined;
}

/** 渐进解析器向大纲 Provider 追加批次时使用的写入接口。 */
export interface ProgressiveWordOutlineProvider extends WordOutlineProvider {
  append(items: readonly WordOutlineItem[]): void;
  complete(): void;
}
