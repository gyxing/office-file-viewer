import { Tree } from 'antd';
import type { Key, MutableRefObject, RefObject } from 'react';
import React, {
  memo,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
} from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type { WordOutlineProvider } from '../../services/word/WordOutlineProvider';
import type { WordPageSource } from '../../services/word/WordPageSource';
import type {
  WordOutlineItem,
  WordOutlineTreeNode,
} from '../../services/word/types';
import {
  OutlineIcon,
  PanelLeftCloseIcon,
  PanelLeftOpenIcon,
} from '../../shell/icons';
import type { WordBlockPageIndex } from '../word-pages/WordBlockPageIndex';
import type { WordPageNavigationController } from '../word-pages/types';
import './index.less';
import { useWordOutlineNavigation } from './useWordOutlineNavigation';

type WordOutlineSidebarProps = {
  /** 源文档明确声明的大纲条目，正文定位继续使用同一稳定 ID。 */
  items: WordOutlineItem[];
  /** 支持普通完整树和大文件按层读取的大纲 Provider。 */
  provider: WordOutlineProvider;
  /** 当前画像选择的 Tree 渲染路径。 */
  outlineMode: 'normal' | 'virtual';
  /** 当前正文选择的完整或窗口页面路径。 */
  pageMode: 'normal' | 'windowed';
  /** 窗口模式用于按目标块优先分页。 */
  pageSource?: WordPageSource<unknown>;
  /** 窗口模式使用的块—页 revision 索引。 */
  blockPageIndex?: WordBlockPageIndex;
  /** 窗口页面挂载和滚动控制器。 */
  pageNavigationControllerRef?: MutableRefObject<
    WordPageNavigationController | undefined
  >;
  /** Word 正文滚动容器。 */
  scrollContainerRef: RefObject<HTMLElement>;
  /** 用于在切换文件时重置侧栏状态。 */
  documentSessionId: string;
  /** 缩放、分页等会改变正文位置的布局键。 */
  layoutKey: string;
};

type OutlineTreeProps = {
  treeData: WordOutlineTreeNode[];
  expandedKeys: Key[];
  selectedKeys: Key[];
  treeLabel: string;
  onExpand(keys: Key[]): void;
  onSelect(keys: Key[]): void;
  titleRender(node: WordOutlineTreeNode): React.ReactNode;
};

/** 收集当前已物化树的全部节点键，作为普通路径的默认展开集合。 */
function collectTreeKeys(nodes: readonly WordOutlineTreeNode[]): string[] {
  return nodes.flatMap((node) => [
    node.key,
    ...collectTreeKeys(node.children ?? []),
  ]);
}

/** 将按需读取的直接子级安装到当前虚拟树，避免重建已展开分支。 */
function installTreeChildren(
  nodes: readonly WordOutlineTreeNode[],
  nodeId: string,
  children: WordOutlineTreeNode[],
): WordOutlineTreeNode[] {
  return nodes.map((node) => {
    if (node.key === nodeId) return { ...node, children };
    if (!node.children?.length) return node;
    const nextChildren = installTreeChildren(node.children, nodeId, children);
    return nextChildren === node.children
      ? node
      : { ...node, children: nextChildren };
  });
}

/** 普通大纲固定关闭虚拟滚动，不测量容器高度。 */
function NormalOutlineTree({
  treeData,
  expandedKeys,
  selectedKeys,
  treeLabel,
  onExpand,
  onSelect,
  titleRender,
}: OutlineTreeProps) {
  return (
    <Tree<WordOutlineTreeNode>
      className="office-file-word-outline__tree"
      aria-label={treeLabel}
      blockNode
      showLine={{ showLeafIcon: false }}
      virtual={false}
      treeData={treeData}
      expandedKeys={expandedKeys}
      selectedKeys={selectedKeys}
      onExpand={onExpand}
      onSelect={(keys, info) => onSelect(keys.length ? keys : [info.node.key])}
      titleRender={titleRender}
    />
  );
}

type VirtualOutlineTreeProps = OutlineTreeProps & {
  provider: WordOutlineProvider;
  documentSessionId: string;
  revision: number;
};

/** 大纲超过阈值时按容器高度虚拟化，并逐层读取子节点。 */
function VirtualOutlineTree({
  provider,
  documentSessionId,
  revision,
  expandedKeys,
  selectedKeys,
  treeLabel,
  onExpand,
  onSelect,
  titleRender,
}: VirtualOutlineTreeProps) {
  const hostRef = useRef<HTMLDivElement>(null);
  const [treeHeight, setTreeHeight] = useState(240);
  const [treeData, setTreeData] = useState<WordOutlineTreeNode[]>(() =>
    provider.getRoots(2),
  );
  const [loadedKeys, setLoadedKeys] = useState<Key[]>([]);

  useEffect(() => {
    setTreeData(provider.getRoots(2));
    setLoadedKeys([]);
  }, [documentSessionId, provider, revision]);

  useLayoutEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const updateHeight = () => {
      const nextHeight = Math.max(120, Math.floor(host.clientHeight));
      setTreeHeight((current) =>
        current === nextHeight ? current : nextHeight,
      );
    };
    updateHeight();
    if (typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(updateHeight);
    observer.observe(host);
    return () => observer.disconnect();
  }, [documentSessionId]);

  return (
    <div
      ref={hostRef}
      className="office-file-word-outline__tree-host office-file-word-outline__tree-host--virtual"
    >
      <Tree<WordOutlineTreeNode>
        className="office-file-word-outline__tree"
        aria-label={treeLabel}
        blockNode
        showLine={{ showLeafIcon: false }}
        virtual
        height={treeHeight}
        treeData={treeData}
        expandedKeys={expandedKeys}
        loadedKeys={loadedKeys}
        selectedKeys={selectedKeys}
        loadData={async (node) => {
          const nodeId = String(node.key);
          const children = await provider.getChildren(nodeId);
          setTreeData((current) =>
            installTreeChildren(current, nodeId, children),
          );
          setLoadedKeys((current) =>
            current.includes(node.key) ? current : [...current, node.key],
          );
        }}
        onExpand={onExpand}
        onSelect={(keys, info) =>
          onSelect(keys.length ? keys : [info.node.key])
        }
        titleRender={titleRender}
      />
    </div>
  );
}

/** 渲染 Word 文档大纲侧栏。 */
function WordOutlineSidebarComponent({
  items,
  provider,
  outlineMode,
  pageMode,
  pageSource,
  blockPageIndex,
  pageNavigationControllerRef,
  scrollContainerRef,
  documentSessionId,
  layoutKey,
}: WordOutlineSidebarProps) {
  const messages = useOfficeFileViewerMessages();
  const snapshot = useSyncExternalStore(
    provider.subscribe,
    provider.getSnapshot,
    provider.getSnapshot,
  );
  const normalTree = useMemo(
    () => (outlineMode === 'normal' ? provider.getRoots() : []),
    [outlineMode, provider, snapshot.revision],
  );
  const initialExpandedKeys = useMemo(
    () =>
      outlineMode === 'normal'
        ? collectTreeKeys(normalTree)
        : provider.getRoots(1).map((node) => node.key),
    [normalTree, outlineMode, provider, snapshot.revision],
  );
  const [collapsed, setCollapsed] = useState(false);
  const [expandedKeys, setExpandedKeys] = useState<Key[]>(initialExpandedKeys);
  const { activeKey, selectTarget } = useWordOutlineNavigation({
    items,
    scrollContainerRef,
    layoutKey,
    enabled: !collapsed && snapshot.count > 0,
    pageMode,
    pageSource,
    blockPageIndex,
    pageNavigationControllerRef,
    documentSessionId,
  });

  useEffect(() => {
    setCollapsed(false);
    setExpandedKeys(initialExpandedKeys);
  }, [documentSessionId, initialExpandedKeys, outlineMode]);

  useEffect(() => {
    if (!activeKey) return;
    let cancelled = false;
    void provider.getAncestors(activeKey).then((ancestors) => {
      if (cancelled || !ancestors.length) return;
      setExpandedKeys((current) => {
        const next = new Set(current);
        const previousSize = next.size;
        ancestors.forEach((key) => next.add(key));
        return next.size === previousSize ? current : [...next];
      });
    });
    return () => {
      cancelled = true;
    };
  }, [activeKey, provider, snapshot.revision]);

  if (!snapshot.count) return null;

  const handleSelect = (keys: Key[]) => {
    const selected = keys[0] ? provider.locate(String(keys[0])) : undefined;
    if (selected) selectTarget(selected);
  };
  const titleRender = (node: WordOutlineTreeNode) => (
    <span className="office-file-word-outline__node-title" title={node.text}>
      {node.text}
    </span>
  );

  return (
    <aside
      className={`office-file-word-outline${
        collapsed ? ' office-file-word-outline--collapsed' : ''
      }`}
      aria-label={messages.outline.region}
      data-outline-count={snapshot.count}
      data-outline-mode={outlineMode}
    >
      {collapsed ? (
        <div className="office-file-word-outline__rail">
          <button
            type="button"
            className="office-file-word-outline__toggle"
            aria-label={messages.outline.expand}
            title={messages.outline.expand}
            onClick={() => setCollapsed(false)}
          >
            <PanelLeftOpenIcon />
          </button>
          <OutlineIcon className="office-file-word-outline__rail-icon" />
        </div>
      ) : (
        <>
          <header className="office-file-word-outline__header">
            <span className="office-file-word-outline__title">
              <OutlineIcon />
              {messages.outline.title}
            </span>
            <span className="office-file-word-outline__count">
              {snapshot.count}
            </span>
            <button
              type="button"
              className="office-file-word-outline__toggle"
              aria-label={messages.outline.collapse}
              title={messages.outline.collapse}
              onClick={() => setCollapsed(true)}
            >
              <PanelLeftCloseIcon />
            </button>
          </header>
          {outlineMode === 'virtual' ? (
            <VirtualOutlineTree
              provider={provider}
              documentSessionId={documentSessionId}
              revision={snapshot.revision}
              treeData={[]}
              expandedKeys={expandedKeys}
              selectedKeys={activeKey ? [activeKey] : []}
              treeLabel={messages.outline.tree}
              onExpand={setExpandedKeys}
              onSelect={handleSelect}
              titleRender={titleRender}
            />
          ) : (
            <div className="office-file-word-outline__tree-host">
              <NormalOutlineTree
                treeData={normalTree}
                expandedKeys={expandedKeys}
                selectedKeys={activeKey ? [activeKey] : []}
                treeLabel={messages.outline.tree}
                onExpand={setExpandedKeys}
                onSelect={handleSelect}
                titleRender={titleRender}
              />
            </div>
          )}
        </>
      )}
    </aside>
  );
}

export const WordOutlineSidebar = memo(WordOutlineSidebarComponent);
