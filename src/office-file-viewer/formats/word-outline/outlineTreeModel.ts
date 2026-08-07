import type { WordOutlineTreeNode } from '../../services/word/types';

/** 自研大纲树渲染一行所需的扁平数据。 */
export type WordOutlineTreeRow = {
  /** 当前行对应的源大纲节点。 */
  node: WordOutlineTreeNode;
  /** 当前节点的稳定标识。 */
  key: string;
  /** 当前节点从零开始的可视层级。 */
  depth: number;
  /** 当前节点的直接父级标识。 */
  parentKey?: string;
  /** 当前节点在同级节点中的索引。 */
  siblingIndex: number;
  /** 当前节点所在层级的节点数量。 */
  siblingCount: number;
  /** 当前节点之后是否还有同级节点。 */
  hasNextSibling: boolean;
  /** 各祖先层级之后是否还有同级分支，用于绘制连续连接线。 */
  ancestorHasNextSibling: boolean[];
};

/** 收集当前已物化树的全部节点键。 */
export function collectOutlineTreeKeys(
  nodes: readonly WordOutlineTreeNode[],
): string[] {
  return nodes.flatMap((node) => [
    node.key,
    ...collectOutlineTreeKeys(node.children ?? []),
  ]);
}

/** 收集已经具备直接子级数据的节点，避免重复发起懒加载请求。 */
export function collectLoadedOutlineTreeKeys(
  nodes: readonly WordOutlineTreeNode[],
  keys = new Set<string>(),
): Set<string> {
  nodes.forEach((node) => {
    if (node.children !== undefined) keys.add(node.key);
    collectLoadedOutlineTreeKeys(node.children ?? [], keys);
  });
  return keys;
}

/** 将按需读取的直接子级安装到当前树，并保留已经物化的后代。 */
export function installOutlineTreeChildren(
  nodes: readonly WordOutlineTreeNode[],
  nodeId: string,
  children: WordOutlineTreeNode[],
): WordOutlineTreeNode[] {
  let changed = false;
  const nextNodes = nodes.map((node) => {
    if (node.key === nodeId) {
      const currentChildren = new Map(
        (node.children ?? []).map((child) => [child.key, child]),
      );
      const nextChildren = children.map((child) => {
        const currentChild = currentChildren.get(child.key);
        return currentChild?.children === undefined
          ? child
          : { ...child, children: currentChild.children };
      });
      changed = true;
      return { ...node, children: nextChildren };
    }
    if (!node.children?.length) return node;
    const nextChildren = installOutlineTreeChildren(
      node.children,
      nodeId,
      children,
    );
    if (nextChildren === node.children) return node;
    changed = true;
    return { ...node, children: nextChildren };
  });
  return changed ? nextNodes : (nodes as WordOutlineTreeNode[]);
}

/** 判断指定节点是否已经物化到当前树中。 */
export function containsOutlineTreeNode(
  nodes: readonly WordOutlineTreeNode[],
  nodeId: string,
): boolean {
  return nodes.some(
    (node) =>
      node.key === nodeId ||
      containsOutlineTreeNode(node.children ?? [], nodeId),
  );
}

/** 将展开后的可见节点转换为固定行高列表，供普通与虚拟路径共用。 */
export function flattenOutlineTree(
  nodes: readonly WordOutlineTreeNode[],
  expandedKeys: ReadonlySet<string>,
): WordOutlineTreeRow[] {
  const rows: WordOutlineTreeRow[] = [];

  const visit = (
    siblings: readonly WordOutlineTreeNode[],
    depth: number,
    parentKey: string | undefined,
    ancestorHasNextSibling: boolean[],
  ) => {
    siblings.forEach((node, siblingIndex) => {
      const hasNextSibling = siblingIndex < siblings.length - 1;
      rows.push({
        node,
        key: node.key,
        depth,
        parentKey,
        siblingIndex,
        siblingCount: siblings.length,
        hasNextSibling,
        ancestorHasNextSibling,
      });
      if (!expandedKeys.has(node.key) || !node.children?.length) return;
      visit(node.children, depth + 1, node.key, [
        ...ancestorHasNextSibling,
        hasNextSibling,
      ]);
    });
  };

  visit(nodes, 0, undefined, []);
  return rows;
}
