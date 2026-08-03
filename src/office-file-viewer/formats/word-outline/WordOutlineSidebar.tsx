import { Tree } from 'antd';
import type { Key, MutableRefObject, RefObject } from 'react';
import React, {
  memo,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type { WordOutlineProvider } from '../../services/word/WordOutlineProvider';
import type { WordPageSource } from '../../services/word/WordPageSource';
import type {
  WordOutlineItem,
  WordOutlineTreeNode,
} from '../../services/word/types';
import { useExternalStoreSnapshot } from '../../shared/react/useExternalStoreSnapshot';
import { OutlineIcon, PanelLeftCloseIcon } from '../../shell/icons';
import type { WordBlockPageIndex } from '../word-pages/WordBlockPageIndex';
import type { WordPageNavigationController } from '../word-pages/types';
import './index.less';
import { useWordOutlineNavigation } from './useWordOutlineNavigation';
import { useWordOutlineResize } from './useWordOutlineResize';

/** Word 大纲侧栏组件属性。 */
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
  /** 关闭大纲侧栏。 */
  onClose: () => void;
};

/** Word 大纲树组件属性。 */
type OutlineTreeProps = {
  /** Ant Design Tree 使用的大纲节点数据。 */
  treeData: WordOutlineTreeNode[];
  /** 当前展开的大纲节点键。 */
  expandedKeys: Key[];
  /** 当前选中的大纲节点键。 */
  selectedKeys: Key[];
  /** 大纲树的无障碍名称。 */
  treeLabel: string;
  /** 在大纲节点展开状态变化时触发。 */
  onExpand(keys: Key[]): void;
  /** 在用户选择大纲节点时触发。 */
  onSelect(keys: Key[]): void;
  /** 生成大纲树节点的显示内容。 */
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

/** Word 大纲虚拟树组件属性。 */
type VirtualOutlineTreeProps = OutlineTreeProps & {
  /** 按需提供当前数据的接口。 */
  provider: WordOutlineProvider;
  /** 当前文档解析会话的标识。 */
  documentSessionId: string;
  /** 数据源变更时递增的修订号。 */
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

/** Word 大纲宽度调整分隔条属性。 */
type WordOutlineResizeHandleProps = {
  /** 待调整宽度的侧栏元素。 */
  panelRef: RefObject<HTMLElement>;
  /** 用于在切换文件时恢复默认宽度。 */
  documentSessionId: string;
  /** 分隔条的无障碍名称。 */
  label: string;
};

/** 渲染支持指针与键盘操作的大纲宽度分隔条。 */
function WordOutlineResizeHandleComponent({
  panelRef,
  documentSessionId,
  label,
}: WordOutlineResizeHandleProps) {
  const {
    width,
    maxWidth,
    minWidth,
    handleRef,
    handleKeyDown,
    handlePointerDown,
    handlePointerMove,
    handlePointerEnd,
  } = useWordOutlineResize(panelRef, documentSessionId);

  return (
    <div
      ref={handleRef}
      className="office-file-word-outline__resize-handle"
      role="separator"
      aria-label={label}
      aria-orientation="vertical"
      aria-valuemin={minWidth}
      aria-valuemax={Math.round(maxWidth)}
      aria-valuenow={Math.round(width)}
      title={label}
      tabIndex={0}
      onKeyDown={handleKeyDown}
      onPointerDown={handlePointerDown}
      onPointerMove={handlePointerMove}
      onPointerUp={handlePointerEnd}
      onPointerCancel={handlePointerEnd}
      onLostPointerCapture={handlePointerEnd}
    >
      <span aria-hidden="true" />
    </div>
  );
}

const WordOutlineResizeHandle = memo(WordOutlineResizeHandleComponent);

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
  onClose,
}: WordOutlineSidebarProps) {
  const messages = useOfficeFileViewerMessages();
  const panelRef = useRef<HTMLElement>(null);
  const snapshot = useExternalStoreSnapshot(provider);
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
  const [expandedKeys, setExpandedKeys] = useState<Key[]>(initialExpandedKeys);
  const { activeKey, selectTarget } = useWordOutlineNavigation({
    items,
    scrollContainerRef,
    layoutKey,
    enabled: snapshot.count > 0,
    pageMode,
    pageSource,
    blockPageIndex,
    pageNavigationControllerRef,
    documentSessionId,
  });

  useEffect(() => {
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
      ref={panelRef}
      className="office-file-word-outline"
      aria-label={messages.outline.region}
      data-outline-count={snapshot.count}
      data-outline-mode={outlineMode}
    >
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
          onClick={onClose}
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
      <WordOutlineResizeHandle
        panelRef={panelRef}
        documentSessionId={documentSessionId}
        label={messages.outline.resize}
      />
    </aside>
  );
}

export const WordOutlineSidebar = memo(WordOutlineSidebarComponent);
