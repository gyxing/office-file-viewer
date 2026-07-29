import type {
  WordOutlineItem,
  WordOutlineTreeNode,
} from '../../services/word/types';

/**
 * 将源文档的扁平大纲转换为树；缺失中间级别时挂到最近的较浅级别，
 * 避免为不完整但合法的 Word 大纲制造虚拟节点。
 */
export function buildWordOutlineTree(
  items: WordOutlineItem[],
): WordOutlineTreeNode[] {
  const roots: WordOutlineTreeNode[] = [];
  const stack: WordOutlineTreeNode[] = [];

  items.forEach((item) => {
    while (stack.length && stack[stack.length - 1].level >= item.level) {
      stack.pop();
    }

    const node: WordOutlineTreeNode = {
      ...item,
      key: item.id,
      children: [],
    };
    const parent = stack[stack.length - 1];
    if (parent) parent.children.push(node);
    else roots.push(node);
    stack.push(node);
  });

  return roots;
}

/** 生成每个大纲节点的祖先键索引，供定位当前项时自动展开父级。 */
export function collectOutlineAncestorKeys(
  nodes: WordOutlineTreeNode[],
): Record<string, string[]> {
  const ancestorsByKey: Record<string, string[]> = {};

  const visit = (node: WordOutlineTreeNode, ancestors: string[]) => {
    ancestorsByKey[node.key] = ancestors;
    const nextAncestors = [...ancestors, node.key];
    node.children.forEach((child) => visit(child, nextAncestors));
  };

  nodes.forEach((node) => visit(node, []));
  return ancestorsByKey;
}
