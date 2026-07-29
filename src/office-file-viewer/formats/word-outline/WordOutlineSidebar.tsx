import { Tree } from 'antd';
import type { Key, RefObject } from 'react';
import React, {
  memo,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import type {
  WordOutlineItem,
  WordOutlineTreeNode,
} from '../../services/word/types';
import {
  OutlineIcon,
  PanelLeftCloseIcon,
  PanelLeftOpenIcon,
} from '../../shell/icons';
import {
  buildWordOutlineTree,
  collectOutlineAncestorKeys,
} from './buildWordOutlineTree';
import './index.less';
import { useWordOutlineNavigation } from './useWordOutlineNavigation';

type WordOutlineSidebarProps = {
  /** 源文档明确声明的大纲条目。 */
  items: WordOutlineItem[];
  /** Word 正文滚动容器。 */
  scrollContainerRef: RefObject<HTMLElement>;
  /** 用于在切换文件时重置侧栏状态。 */
  documentIdentity: object;
  /** 缩放、分页等会改变正文位置的布局键。 */
  layoutKey: string;
};

/** 收集树的全部节点键，作为默认展开集合。 */
function collectTreeKeys(nodes: WordOutlineTreeNode[]): string[] {
  return nodes.flatMap((node) => [node.key, ...collectTreeKeys(node.children)]);
}

/** 渲染 Word 文档大纲侧栏。 */
function WordOutlineSidebarComponent({
  items,
  scrollContainerRef,
  documentIdentity,
  layoutKey,
}: WordOutlineSidebarProps) {
  const tree = useMemo(() => buildWordOutlineTree(items), [items]);
  const allKeys = useMemo(() => collectTreeKeys(tree), [tree]);
  const ancestorsByKey = useMemo(
    () => collectOutlineAncestorKeys(tree),
    [tree],
  );
  const itemByKey = useMemo(
    () => new Map(items.map((item) => [item.id, item])),
    [items],
  );
  const [collapsed, setCollapsed] = useState(false);
  const [expandedKeys, setExpandedKeys] = useState<Key[]>(allKeys);
  const [treeHeight, setTreeHeight] = useState(240);
  const treeHostRef = useRef<HTMLDivElement>(null);
  const { activeKey, selectTarget } = useWordOutlineNavigation({
    items,
    scrollContainerRef,
    layoutKey,
    enabled: !collapsed && items.length > 0,
  });

  useEffect(() => {
    setCollapsed(false);
    setExpandedKeys(allKeys);
  }, [allKeys, documentIdentity]);

  useEffect(() => {
    if (!activeKey) return;
    const ancestors = ancestorsByKey[activeKey] ?? [];
    if (!ancestors.length) return;
    setExpandedKeys((current) => {
      const next = new Set(current);
      const previousSize = next.size;
      ancestors.forEach((key) => next.add(key));
      return next.size === previousSize ? current : [...next];
    });
  }, [activeKey, ancestorsByKey]);

  useLayoutEffect(() => {
    const host = treeHostRef.current;
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
  }, [collapsed]);

  if (!items.length) return null;

  return (
    <aside
      className={`office-file-word-outline${
        collapsed ? ' office-file-word-outline--collapsed' : ''
      }`}
      aria-label="文档大纲"
      data-outline-count={items.length}
    >
      {collapsed ? (
        <div className="office-file-word-outline__rail">
          <button
            type="button"
            className="office-file-word-outline__toggle"
            aria-label="展开文档大纲"
            title="展开文档大纲"
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
              大纲
            </span>
            <span className="office-file-word-outline__count">
              {items.length}
            </span>
            <button
              type="button"
              className="office-file-word-outline__toggle"
              aria-label="收起文档大纲"
              title="收起文档大纲"
              onClick={() => setCollapsed(true)}
            >
              <PanelLeftCloseIcon />
            </button>
          </header>
          <div
            ref={treeHostRef}
            className="office-file-word-outline__tree-host"
          >
            <Tree<WordOutlineTreeNode>
              aria-label="大纲目录"
              blockNode
              showLine={{ showLeafIcon: false }}
              virtual
              height={treeHeight}
              treeData={tree}
              expandedKeys={expandedKeys}
              selectedKeys={activeKey ? [activeKey] : []}
              onExpand={(keys) => setExpandedKeys(keys)}
              onSelect={(keys) => {
                const selected = keys[0]
                  ? itemByKey.get(String(keys[0]))
                  : undefined;
                if (selected) selectTarget(selected);
              }}
              titleRender={(node) => (
                <span
                  className="office-file-word-outline__node-title"
                  title={node.text}
                >
                  {node.text}
                </span>
              )}
            />
          </div>
        </>
      )}
    </aside>
  );
}

export const WordOutlineSidebar = memo(WordOutlineSidebarComponent);
