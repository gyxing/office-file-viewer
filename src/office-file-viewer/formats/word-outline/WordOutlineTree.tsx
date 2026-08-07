import React, { memo, useCallback, useEffect, useRef, useState } from 'react';
import type { WordOutlineProvider } from '../../services/word/WordOutlineProvider';
import type { WordOutlineTreeNode } from '../../services/word/types';
import { OfficeOutlineTree } from './OfficeOutlineTree';
import {
  collectLoadedOutlineTreeKeys,
  installOutlineTreeChildren,
} from './outlineTreeModel';

/** 大纲树数据读取和渲染组件属性。 */
type WordOutlineTreeProps = {
  /** 普通模式已经完整物化的大纲树。 */
  normalTreeData: WordOutlineTreeNode[];
  /** 大文件模式使用的按层数据源。 */
  provider: WordOutlineProvider;
  /** 当前是否使用虚拟窗口和懒加载。 */
  virtual: boolean;
  /** 当前文档解析会话的标识。 */
  documentSessionId: string;
  /** 数据源变更时递增的修订号。 */
  revision: number;
  /** 当前展开的节点键。 */
  expandedKeys: readonly string[];
  /** 当前选中的节点键。 */
  selectedKey?: string;
  /** 正文滚动后需要显示到舒适区的节点键。 */
  revealKey?: string;
  /** 大纲树的无障碍名称。 */
  label: string;
  /** 展开节点按钮的无障碍文案。 */
  expandLabel: string;
  /** 收起节点按钮的无障碍文案。 */
  collapseLabel: string;
  /** 切换节点展开状态。 */
  onToggle(key: string, expanded: boolean): void;
  /** 用户选择节点时调用。 */
  onSelect(key: string): void;
};

/** 统一普通完整树与大文件按层读取路径。 */
function WordOutlineTreeComponent({
  normalTreeData,
  provider,
  virtual,
  documentSessionId,
  revision,
  expandedKeys,
  selectedKey,
  revealKey,
  label,
  expandLabel,
  collapseLabel,
  onToggle,
  onSelect,
}: WordOutlineTreeProps) {
  const requestGenerationRef = useRef(0);
  const loadedKeySetRef = useRef(new Set<string>());
  const childRequestRef = useRef(
    new Map<string, Promise<WordOutlineTreeNode[]>>(),
  );
  const [virtualTreeData, setVirtualTreeData] = useState<WordOutlineTreeNode[]>(
    () => provider.getRoots(2),
  );
  const [loadingKeys, setLoadingKeys] = useState<Set<string>>(new Set());

  useEffect(() => {
    requestGenerationRef.current += 1;
    childRequestRef.current.clear();
    const roots = provider.getRoots(2);
    loadedKeySetRef.current = collectLoadedOutlineTreeKeys(roots);
    setVirtualTreeData(roots);
    setLoadingKeys(new Set());
  }, [documentSessionId, provider, revision, virtual]);

  const loadChildren = useCallback(
    async (nodeId: string, expectedGeneration: number) => {
      if (loadedKeySetRef.current.has(nodeId)) return;
      let request = childRequestRef.current.get(nodeId);
      if (!request) {
        request = provider.getChildren(nodeId);
        childRequestRef.current.set(nodeId, request);
        setLoadingKeys((current) => {
          if (current.has(nodeId)) return current;
          const next = new Set(current);
          next.add(nodeId);
          return next;
        });
      }

      try {
        const children = await request;
        if (requestGenerationRef.current !== expectedGeneration) return;
        setVirtualTreeData((current) =>
          installOutlineTreeChildren(current, nodeId, children),
        );
        loadedKeySetRef.current.add(nodeId);
      } catch {
        // 子级读取失败时保持节点可重试；解析错误由主预览状态统一呈现。
      } finally {
        if (requestGenerationRef.current === expectedGeneration) {
          childRequestRef.current.delete(nodeId);
          setLoadingKeys((current) => {
            if (!current.has(nodeId)) return current;
            const next = new Set(current);
            next.delete(nodeId);
            return next;
          });
        }
      }
    },
    [provider],
  );

  useEffect(() => {
    if (!virtual || !revealKey) return;
    let cancelled = false;
    const expectedGeneration = requestGenerationRef.current;

    void (async () => {
      const ancestorKeys = await provider.getAncestors(revealKey);
      if (cancelled || requestGenerationRef.current !== expectedGeneration) {
        return;
      }
      for (const ancestorKey of ancestorKeys) {
        if (cancelled) return;
        await loadChildren(ancestorKey, expectedGeneration);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [loadChildren, provider, revealKey, revision, virtual]);

  const handleToggle = (key: string, expanded: boolean) => {
    onToggle(key, expanded);
    if (!virtual || !expanded || loadedKeySetRef.current.has(key)) return;
    void loadChildren(key, requestGenerationRef.current);
  };

  return (
    <OfficeOutlineTree
      treeData={virtual ? virtualTreeData : normalTreeData}
      expandedKeys={expandedKeys}
      selectedKey={selectedKey}
      revealKey={revealKey}
      virtual={virtual}
      label={label}
      expandLabel={expandLabel}
      collapseLabel={collapseLabel}
      loadingKeys={loadingKeys}
      onToggle={handleToggle}
      onSelect={onSelect}
    />
  );
}

export const WordOutlineTree = memo(WordOutlineTreeComponent);
