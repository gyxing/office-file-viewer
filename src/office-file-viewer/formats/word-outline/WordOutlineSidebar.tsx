import { Tree } from 'antd';
import type { ComponentRef, Key, MutableRefObject, RefObject } from 'react';
import React, {
  memo,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type {
  WordOutlineProvider,
  WordOutlineProviderSnapshot,
} from '../../services/word/WordOutlineProvider';
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
  /** 当前是否向用户显示大纲侧栏。 */
  visible: boolean;
  /** 当前文档的大纲是否已经完成首次激活。 */
  activated: boolean;
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
  /** 由正文滚动触发、需要自动显示到舒适区的节点键。 */
  revealKey?: string;
  /** 大纲树的无障碍名称。 */
  treeLabel: string;
  /** 在大纲节点展开状态变化时触发。 */
  onExpand(keys: Key[]): void;
  /** 在用户选择大纲节点时触发。 */
  onSelect(keys: Key[]): void;
  /** 生成大纲树节点的显示内容。 */
  titleRender(node: WordOutlineTreeNode): React.ReactNode;
};

/** 选中项位于该垂直区间时保持目录位置稳定，避免正文滚动时频繁跳动。 */
const OUTLINE_REVEAL_ZONE = {
  start: 0.25,
  end: 0.6,
  target: 0.4,
} as const;

/** 大纲节点的基准行高，与节点标题样式保持一致。 */
const OUTLINE_TREE_ROW_HEIGHT = 28;

/** 未首次展开时不订阅渐进式大纲，避免隐藏侧栏参与解析期重渲染。 */
const INACTIVE_OUTLINE_SNAPSHOT: WordOutlineProviderSnapshot = {
  revision: 0,
  count: 0,
  complete: true,
};

/** 虚拟树等待滚动定位的节点及其祖先路径。 */
type PendingOutlineReveal = {
  /** 需要滚动到可视区域的节点键。 */
  key: string;
  /** 节点从根到父节点的祖先键。 */
  ancestorKeys: string[];
};

/** 收集当前已物化树的全部节点键，作为普通路径的默认展开集合。 */
function collectTreeKeys(nodes: readonly WordOutlineTreeNode[]): string[] {
  return nodes.flatMap((node) => [
    node.key,
    ...collectTreeKeys(node.children ?? []),
  ]);
}

/** 将按需读取的直接子级安装到当前虚拟树，并保留已经物化的后代。 */
function installTreeChildren(
  nodes: readonly WordOutlineTreeNode[],
  nodeId: string,
  children: WordOutlineTreeNode[],
): WordOutlineTreeNode[] {
  return nodes.map((node) => {
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
      return { ...node, children: nextChildren };
    }
    if (!node.children?.length) return node;
    const nextChildren = installTreeChildren(node.children, nodeId, children);
    return nextChildren === node.children
      ? node
      : { ...node, children: nextChildren };
  });
}

/** 判断指定节点是否已经物化到当前虚拟树数据中。 */
function containsTreeNode(
  nodes: readonly WordOutlineTreeNode[],
  nodeId: string,
): boolean {
  return nodes.some(
    (node) =>
      node.key === nodeId || containsTreeNode(node.children ?? [], nodeId),
  );
}

/** 查找当前已经渲染到 DOM 中的大纲节点。 */
function findRenderedOutlineNode(
  host: HTMLElement,
  selectedKey: string,
): HTMLElement | undefined {
  return Array.from(
    host.querySelectorAll<HTMLElement>('[data-office-file-word-outline-key]'),
  ).find((node) => node.dataset.officeFileWordOutlineKey === selectedKey);
}

/** 返回将节点中心移动到舒适位置所需的滚动距离，区间内无需滚动。 */
function getOutlineRevealDelta(
  host: HTMLElement,
  selectedNode: HTMLElement,
): number | undefined {
  const hostRect = host.getBoundingClientRect();
  const selectedRect = selectedNode.getBoundingClientRect();
  const selectedCenter = (selectedRect.top + selectedRect.bottom) / 2;
  const relativeCenter = selectedCenter - hostRect.top;
  const comfortStart = host.clientHeight * OUTLINE_REVEAL_ZONE.start;
  const comfortEnd = host.clientHeight * OUTLINE_REVEAL_ZONE.end;
  if (relativeCenter >= comfortStart && relativeCenter <= comfortEnd) {
    return undefined;
  }
  return relativeCenter - host.clientHeight * OUTLINE_REVEAL_ZONE.target;
}

/** 普通大纲关闭虚拟滚动，并在 DOM 容器内维护舒适区定位。 */
function NormalOutlineTree({
  treeData,
  expandedKeys,
  selectedKeys,
  revealKey,
  treeLabel,
  onExpand,
  onSelect,
  titleRender,
}: OutlineTreeProps) {
  const hostRef = useRef<HTMLDivElement>(null);

  useLayoutEffect(() => {
    const host = hostRef.current;
    if (!host || !revealKey) return;
    const selectedNode = findRenderedOutlineNode(host, revealKey);
    if (!selectedNode) return;
    const revealDelta = getOutlineRevealDelta(host, selectedNode);
    if (revealDelta === undefined) return;
    const maxScrollTop = Math.max(0, host.scrollHeight - host.clientHeight);
    host.scrollTop = Math.max(
      0,
      Math.min(maxScrollTop, host.scrollTop + revealDelta),
    );
  }, [expandedKeys, revealKey, treeData]);

  return (
    <div ref={hostRef} className="office-file-word-outline__tree-host">
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
        onSelect={(keys, info) =>
          onSelect(keys.length ? keys : [info.node.key])
        }
        titleRender={titleRender}
      />
    </div>
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
  revealKey,
  treeLabel,
  onExpand,
  onSelect,
  titleRender,
}: VirtualOutlineTreeProps) {
  const hostRef = useRef<HTMLDivElement>(null);
  const treeRef = useRef<ComponentRef<typeof Tree>>(null);
  const loadedKeySetRef = useRef(new Set<string>());
  const requestGenerationRef = useRef(0);
  const [treeHeight, setTreeHeight] = useState(240);
  const [treeData, setTreeData] = useState<WordOutlineTreeNode[]>(() =>
    provider.getRoots(2),
  );
  const [loadedKeys, setLoadedKeys] = useState<Key[]>([]);
  const [pendingReveal, setPendingReveal] = useState<PendingOutlineReveal>();
  const expandedKeySet = useMemo(
    () => new Set(expandedKeys.map(String)),
    [expandedKeys],
  );

  useEffect(() => {
    requestGenerationRef.current += 1;
    loadedKeySetRef.current.clear();
    setTreeData(provider.getRoots(2));
    setLoadedKeys([]);
    setPendingReveal(undefined);
  }, [documentSessionId, provider, revision]);

  useEffect(() => {
    if (!revealKey) {
      setPendingReveal(undefined);
      return;
    }
    let cancelled = false;
    const requestGeneration = requestGenerationRef.current;

    void (async () => {
      const ancestorKeys = await provider.getAncestors(revealKey);
      const materializedBranches: Array<{
        key: string;
        children: WordOutlineTreeNode[];
      }> = [];
      for (const ancestorKey of ancestorKeys) {
        if (loadedKeySetRef.current.has(ancestorKey)) continue;
        materializedBranches.push({
          key: ancestorKey,
          children: await provider.getChildren(ancestorKey),
        });
      }
      if (cancelled || requestGeneration !== requestGenerationRef.current) {
        return;
      }

      if (materializedBranches.length) {
        setTreeData((current) =>
          materializedBranches.reduce(
            (next, branch) =>
              installTreeChildren(next, branch.key, branch.children),
            current,
          ),
        );
      }
      ancestorKeys.forEach((key) => loadedKeySetRef.current.add(key));
      setLoadedKeys((current) => {
        const next = new Set(current);
        ancestorKeys.forEach((key) => next.add(key));
        return next.size === current.length ? current : [...next];
      });
      setPendingReveal({ key: revealKey, ancestorKeys });
    })();

    return () => {
      cancelled = true;
    };
  }, [documentSessionId, provider, revealKey, revision]);

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

  useLayoutEffect(() => {
    if (!pendingReveal || pendingReveal.key !== revealKey) return;
    if (!containsTreeNode(treeData, pendingReveal.key)) return;
    if (pendingReveal.ancestorKeys.some((key) => !expandedKeySet.has(key))) {
      return;
    }
    const tree = treeRef.current;
    if (!tree) return;

    const host = hostRef.current;
    const renderedNode =
      host && findRenderedOutlineNode(host, pendingReveal.key);
    if (
      host &&
      renderedNode &&
      getOutlineRevealDelta(host, renderedNode) === undefined
    ) {
      setPendingReveal(undefined);
      return;
    }

    let animationFrame: number | undefined;
    const revealSelectedNode = () => {
      tree.scrollTo({
        key: pendingReveal.key,
        align: 'top',
        offset: Math.max(
          0,
          Math.round(
            treeHeight * OUTLINE_REVEAL_ZONE.target -
              OUTLINE_TREE_ROW_HEIGHT / 2,
          ),
        ),
      });
      setPendingReveal((current) =>
        current?.key === pendingReveal.key ? undefined : current,
      );
    };
    if (typeof requestAnimationFrame === 'undefined') {
      revealSelectedNode();
    } else {
      animationFrame = requestAnimationFrame(revealSelectedNode);
    }
    return () => {
      if (animationFrame !== undefined) cancelAnimationFrame(animationFrame);
    };
  }, [expandedKeySet, pendingReveal, revealKey, treeData, treeHeight]);

  return (
    <div
      ref={hostRef}
      className="office-file-word-outline__tree-host office-file-word-outline__tree-host--virtual"
    >
      <Tree<WordOutlineTreeNode>
        ref={treeRef}
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
          const requestGeneration = requestGenerationRef.current;
          const children = await provider.getChildren(nodeId);
          if (requestGeneration !== requestGenerationRef.current) return;
          setTreeData((current) =>
            installTreeChildren(current, nodeId, children),
          );
          loadedKeySetRef.current.add(nodeId);
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
  visible,
  activated,
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
  // 用户点击目录时只滚动正文，避免左侧已可见节点被再次强制定位。
  const manuallySelectedKeyRef = useRef<string>();
  const snapshot = useExternalStoreSnapshot(
    activated ? provider : undefined,
    INACTIVE_OUTLINE_SNAPSHOT,
  );
  const normalTree = useMemo(
    () => (activated && outlineMode === 'normal' ? provider.getRoots() : []),
    [activated, outlineMode, provider, snapshot.revision],
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
    enabled: visible && activated && snapshot.count > 0,
    pageMode,
    pageSource,
    blockPageIndex,
    pageNavigationControllerRef,
    documentSessionId,
  });
  const revealKey =
    activeKey && activeKey !== manuallySelectedKeyRef.current
      ? activeKey
      : undefined;

  useEffect(() => {
    setExpandedKeys(initialExpandedKeys);
  }, [documentSessionId, initialExpandedKeys, outlineMode]);

  useEffect(() => {
    manuallySelectedKeyRef.current = undefined;
  }, [documentSessionId]);

  useEffect(() => {
    const panel = panelRef.current;
    if (!panel) return;
    // 收起期间保留 DOM 供动画使用，但不能让隐藏目录继续接收焦点。
    if (visible) panel.removeAttribute('inert');
    else panel.setAttribute('inert', '');
  }, [visible]);

  useEffect(() => {
    if (
      activeKey &&
      manuallySelectedKeyRef.current &&
      activeKey !== manuallySelectedKeyRef.current
    ) {
      manuallySelectedKeyRef.current = undefined;
    }
  }, [activeKey]);

  useEffect(() => {
    if (!visible || !activeKey) return;
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
  }, [activeKey, provider, snapshot.revision, visible]);

  const handleSelect = (keys: Key[]) => {
    const selected = keys[0] ? provider.locate(String(keys[0])) : undefined;
    if (selected) {
      manuallySelectedKeyRef.current = selected.id;
      selectTarget(selected);
    }
  };
  const titleRender = (node: WordOutlineTreeNode) => (
    <span
      className="office-file-word-outline__node-title"
      data-office-file-word-outline-key={node.key}
      title={node.text}
    >
      {node.text}
    </span>
  );

  return (
    <aside
      ref={panelRef}
      className="office-file-word-outline"
      aria-label={messages.outline.region}
      aria-hidden={!visible}
      data-visible={visible ? 'true' : 'false'}
      data-outline-count={snapshot.count}
      data-outline-mode={outlineMode}
    >
      <div className="office-file-word-outline__viewport">
        <div className="office-file-word-outline__surface">
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
              onClick={(event) => {
                event.currentTarget.blur();
                onClose();
              }}
            >
              <PanelLeftCloseIcon />
            </button>
          </header>
          {activated && snapshot.count > 0 ? (
            outlineMode === 'virtual' ? (
              <VirtualOutlineTree
                provider={provider}
                documentSessionId={documentSessionId}
                revision={snapshot.revision}
                treeData={[]}
                expandedKeys={expandedKeys}
                selectedKeys={activeKey ? [activeKey] : []}
                revealKey={revealKey}
                treeLabel={messages.outline.tree}
                onExpand={setExpandedKeys}
                onSelect={handleSelect}
                titleRender={titleRender}
              />
            ) : (
              <NormalOutlineTree
                treeData={normalTree}
                expandedKeys={expandedKeys}
                selectedKeys={activeKey ? [activeKey] : []}
                revealKey={revealKey}
                treeLabel={messages.outline.tree}
                onExpand={setExpandedKeys}
                onSelect={handleSelect}
                titleRender={titleRender}
              />
            )
          ) : null}
        </div>
      </div>
      <WordOutlineResizeHandle
        panelRef={panelRef}
        documentSessionId={documentSessionId}
        label={messages.outline.resize}
      />
    </aside>
  );
}

export const WordOutlineSidebar = memo(WordOutlineSidebarComponent);
