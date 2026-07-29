import type {
  ProgressiveWordOutlineProvider,
  WordOutlineProvider,
  WordOutlineProviderSnapshot,
} from './WordOutlineProvider';
import type { WordOutlineItem, WordOutlineTreeNode } from './types';

type IndexedOutlineNode = {
  item: WordOutlineItem;
  childIds: string[];
  ancestorIds: string[];
};

/** 创建可被普通模型一次填充、也可被解析器分批追加的大纲索引。 */
function createIndexedWordOutlineProvider(
  initialItems: readonly WordOutlineItem[],
  initiallyComplete: boolean,
): ProgressiveWordOutlineProvider {
  const nodes = new Map<string, IndexedOutlineNode>();
  const rootIds: string[] = [];
  const stack: IndexedOutlineNode[] = [];
  const listeners = new Set<() => void>();
  let snapshot: WordOutlineProviderSnapshot = {
    revision: 0,
    count: 0,
    complete: initiallyComplete,
  };

  const appendItems = (items: readonly WordOutlineItem[]) => {
    let appended = false;
    items.forEach((item) => {
      if (nodes.has(item.id)) return;
      while (stack.length && stack[stack.length - 1].item.level >= item.level) {
        stack.pop();
      }
      const parent = stack[stack.length - 1];
      const node: IndexedOutlineNode = {
        item,
        childIds: [],
        ancestorIds: parent ? [...parent.ancestorIds, parent.item.id] : [],
      };
      nodes.set(item.id, node);
      if (parent) parent.childIds.push(item.id);
      else rootIds.push(item.id);
      stack.push(node);
      appended = true;
    });
    if (!appended) return;
    snapshot = {
      ...snapshot,
      revision: snapshot.revision + 1,
      count: nodes.size,
    };
    listeners.forEach((listener) => listener());
  };

  const materializeNode = (
    nodeId: string,
    currentDepth: number,
    maxDepth?: number,
  ): WordOutlineTreeNode | undefined => {
    const indexed = nodes.get(nodeId);
    if (!indexed) return undefined;
    const mayIncludeChildren =
      maxDepth === undefined || currentDepth < maxDepth;
    const children = mayIncludeChildren
      ? indexed.childIds.flatMap((childId) => {
          const child = materializeNode(childId, currentDepth + 1, maxDepth);
          return child ? [child] : [];
        })
      : undefined;
    return {
      ...indexed.item,
      key: indexed.item.id,
      isLeaf: indexed.childIds.length === 0,
      children,
    };
  };

  appendItems(initialItems);
  if (initiallyComplete) {
    snapshot = { ...snapshot, complete: true };
  }

  return {
    getSnapshot: () => snapshot,
    subscribe: (listener) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    getRoots: (maxDepth) =>
      rootIds.flatMap((rootId) => {
        const root = materializeNode(rootId, 1, maxDepth);
        return root ? [root] : [];
      }),
    getChildren: async (nodeId) => {
      const node = nodes.get(nodeId);
      if (!node) return [];
      return node.childIds.flatMap((childId) => {
        const child = materializeNode(childId, 1, 1);
        return child ? [child] : [];
      });
    },
    getAncestors: async (nodeId) => [...(nodes.get(nodeId)?.ancestorIds ?? [])],
    locate: (nodeId) => nodes.get(nodeId)?.item,
    append: appendItems,
    complete: () => {
      if (snapshot.complete) return;
      snapshot = {
        ...snapshot,
        revision: snapshot.revision + 1,
        complete: true,
      };
      listeners.forEach((listener) => listener());
    },
  };
}

/** 为已完整物化的 Word 模型创建固定大纲 Provider。 */
export function createMemoryWordOutlineProvider(
  items: readonly WordOutlineItem[],
): WordOutlineProvider {
  return createIndexedWordOutlineProvider(items, true);
}

/** 为大文件渐进解析创建可分批追加的大纲 Provider。 */
export function createProgressiveWordOutlineProvider(): ProgressiveWordOutlineProvider {
  return createIndexedWordOutlineProvider([], false);
}
