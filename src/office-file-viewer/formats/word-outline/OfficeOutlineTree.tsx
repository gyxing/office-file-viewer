import type { KeyboardEvent, MouseEvent, UIEvent } from 'react';
import React, {
  memo,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import type { WordOutlineTreeNode } from '../../services/word/types';
import {
  flattenOutlineTree,
  type WordOutlineTreeRow,
} from './outlineTreeModel';

/** 自研大纲树组件属性。 */
type OfficeOutlineTreeProps = {
  /** 当前已经物化的大纲树。 */
  treeData: WordOutlineTreeNode[];
  /** 当前展开的节点键。 */
  expandedKeys: readonly string[];
  /** 当前选中的节点键。 */
  selectedKey?: string;
  /** 正文滚动后需要显示到舒适区的节点键。 */
  revealKey?: string;
  /** 大纲数量较大时启用固定行高虚拟窗口。 */
  virtual: boolean;
  /** 大纲树的无障碍名称。 */
  label: string;
  /** 展开节点按钮的无障碍文案。 */
  expandLabel: string;
  /** 收起节点按钮的无障碍文案。 */
  collapseLabel: string;
  /** 当前正在按需读取子节点的键集合。 */
  loadingKeys: ReadonlySet<string>;
  /** 切换指定节点的展开状态。 */
  onToggle(key: string, expanded: boolean): void;
  /** 用户选择大纲节点时调用。 */
  onSelect(key: string): void;
};

/** 选中项位于该垂直区间时保持目录位置稳定，避免正文滚动时频繁跳动。 */
const OUTLINE_REVEAL_ZONE = {
  start: 0.25,
  end: 0.6,
  target: 0.4,
} as const;

/** 大纲标题固定为单行，普通和虚拟路径共用该行高。 */
const OUTLINE_TREE_ROW_HEIGHT = 28;
/** 虚拟窗口上下额外渲染的行数，降低快速滚动时的空窗概率。 */
const OUTLINE_TREE_OVERSCAN = 8;
/** 树内容顶部与滚动容器之间的留白。 */
const OUTLINE_TREE_TOP_PADDING = 8;

/** 将滚动位置限制在真实内容范围内，避免末尾节点下方出现大块空白。 */
function clampTreeScrollTop(host: HTMLElement, value: number) {
  return Math.max(
    0,
    Math.min(Math.max(0, host.scrollHeight - host.clientHeight), value),
  );
}

/** 渲染经典树状连接线。 */
function OutlineRowGuides({ row }: { row: WordOutlineTreeRow }) {
  if (!row.depth) return null;
  return (
    <span className="office-file-outline-tree__guides" aria-hidden="true">
      {row.ancestorHasNextSibling
        .slice(1)
        .map((continues, level) =>
          continues ? (
            <span
              key={level}
              className="office-file-outline-tree__guide"
              style={{ left: level * 24 + 10 }}
            />
          ) : null,
        )}
      <span
        className="office-file-outline-tree__branch"
        data-continues={row.hasNextSibling ? 'true' : 'false'}
        style={{ left: (row.depth - 1) * 24 + 10 }}
      />
    </span>
  );
}

/** 自研树行属性。 */
type OutlineTreeRowViewProps = {
  /** 当前渲染的扁平树行。 */
  row: WordOutlineTreeRow;
  /** 当前行在完整可见列表中的索引。 */
  rowIndex: number;
  /** 当前节点是否展开。 */
  expanded: boolean;
  /** 当前节点是否选中。 */
  selected: boolean;
  /** 当前节点是否持有漫游焦点。 */
  focused: boolean;
  /** 当前节点是否正在读取子级。 */
  loading: boolean;
  /** 展开节点按钮的无障碍文案。 */
  expandLabel: string;
  /** 收起节点按钮的无障碍文案。 */
  collapseLabel: string;
  /** 保存当前已渲染行的元素。 */
  registerRow(key: string, element: HTMLDivElement | null): void;
  /** 点击节点展开开关时调用。 */
  onToggle(row: WordOutlineTreeRow): void;
  /** 点击或确认节点时调用。 */
  onSelect(row: WordOutlineTreeRow, element: HTMLDivElement): void;
  /** 节点接收焦点时调用。 */
  onFocus(key: string): void;
  /** 节点触发键盘操作时调用。 */
  onKeyDown(event: KeyboardEvent<HTMLDivElement>, rowIndex: number): void;
};

/** 渲染一条具备树语义、连接线和展开开关的大纲行。 */
function OutlineTreeRowView({
  row,
  rowIndex,
  expanded,
  selected,
  focused,
  loading,
  expandLabel,
  collapseLabel,
  registerRow,
  onToggle,
  onSelect,
  onFocus,
  onKeyDown,
}: OutlineTreeRowViewProps) {
  const expandable = !row.node.isLeaf;
  const handleMouseDown = (event: MouseEvent<HTMLDivElement>) => {
    // 浏览器默认聚焦可能自动滚动容器，改由树组件以受控位置聚焦。
    event.preventDefault();
  };

  return (
    <div
      ref={(element) => registerRow(row.key, element)}
      className="office-file-outline-tree__row"
      style={{ paddingLeft: row.depth * 24 }}
      aria-expanded={expandable ? expanded : undefined}
      aria-level={row.depth + 1}
      aria-posinset={row.siblingIndex + 1}
      aria-selected={selected}
      aria-setsize={row.siblingCount}
      data-focused={focused ? 'true' : 'false'}
      data-office-file-word-outline-key={row.key}
      role="treeitem"
      tabIndex={focused ? 0 : -1}
      onClick={(event) => onSelect(row, event.currentTarget)}
      onFocus={() => onFocus(row.key)}
      onKeyDown={(event) => onKeyDown(event, rowIndex)}
      onMouseDown={handleMouseDown}
    >
      <OutlineRowGuides row={row} />
      {expandable ? (
        <button
          className="office-file-outline-tree__switcher"
          type="button"
          aria-label={expanded ? collapseLabel : expandLabel}
          aria-expanded={expanded}
          data-expanded={expanded ? 'true' : 'false'}
          data-loading={loading ? 'true' : 'false'}
          tabIndex={-1}
          onClick={(event) => {
            event.stopPropagation();
            onToggle(row);
          }}
          onMouseDown={(event) => event.preventDefault()}
        >
          <span aria-hidden="true" />
        </button>
      ) : (
        <span className="office-file-outline-tree__leaf-space" />
      )}
      <span className="office-file-outline-tree__title" title={row.node.text}>
        {row.node.text}
      </span>
    </div>
  );
}

/** 渲染普通或按阈值启用虚拟窗口的无依赖大纲树。 */
function OfficeOutlineTreeComponent({
  treeData,
  expandedKeys,
  selectedKey,
  revealKey,
  virtual,
  label,
  expandLabel,
  collapseLabel,
  loadingKeys,
  onToggle,
  onSelect,
}: OfficeOutlineTreeProps) {
  const hostRef = useRef<HTMLDivElement>(null);
  const rowRefs = useRef(new Map<string, HTMLDivElement>());
  const pendingFocusKeyRef = useRef<string>();
  const scrollFrameRef = useRef<number>();
  const latestScrollTopRef = useRef(0);
  const expandedKeySet = useMemo(() => new Set(expandedKeys), [expandedKeys]);
  const rows = useMemo(
    () => flattenOutlineTree(treeData, expandedKeySet),
    [expandedKeySet, treeData],
  );
  const rowIndexByKey = useMemo(
    () => new Map(rows.map((row, index) => [row.key, index])),
    [rows],
  );
  const [focusedKey, setFocusedKey] = useState<string>();
  const [viewportHeight, setViewportHeight] = useState(0);
  const [scrollTop, setScrollTop] = useState(0);

  useLayoutEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const updateHeight = () => {
      const nextHeight = Math.max(0, Math.floor(host.clientHeight));
      setViewportHeight((current) =>
        current === nextHeight ? current : nextHeight,
      );
    };
    updateHeight();
    if (typeof ResizeObserver !== 'undefined') {
      const observer = new ResizeObserver(updateHeight);
      observer.observe(host);
      return () => observer.disconnect();
    }
    window.addEventListener('resize', updateHeight);
    return () => window.removeEventListener('resize', updateHeight);
  }, []);

  useEffect(
    () => () => {
      if (scrollFrameRef.current !== undefined) {
        cancelAnimationFrame(scrollFrameRef.current);
      }
    },
    [],
  );

  useEffect(() => {
    if (focusedKey && rowIndexByKey.has(focusedKey)) return;
    setFocusedKey(
      selectedKey && rowIndexByKey.has(selectedKey)
        ? selectedKey
        : rows[0]?.key,
    );
  }, [focusedKey, rowIndexByKey, rows, selectedKey]);

  useEffect(() => {
    const host = hostRef.current;
    if (
      !selectedKey ||
      !rowIndexByKey.has(selectedKey) ||
      (typeof document !== 'undefined' &&
        host?.contains(document.activeElement))
    ) {
      return;
    }
    setFocusedKey(selectedKey);
  }, [rowIndexByKey, selectedKey]);

  useLayoutEffect(() => {
    const host = hostRef.current;
    const revealIndex = revealKey ? rowIndexByKey.get(revealKey) : undefined;
    if (!host || revealIndex === undefined || !host.clientHeight) return;
    const itemCenter =
      OUTLINE_TREE_TOP_PADDING +
      revealIndex * OUTLINE_TREE_ROW_HEIGHT +
      OUTLINE_TREE_ROW_HEIGHT / 2;
    const relativeCenter = itemCenter - host.scrollTop;
    const comfortStart = host.clientHeight * OUTLINE_REVEAL_ZONE.start;
    const comfortEnd = host.clientHeight * OUTLINE_REVEAL_ZONE.end;
    if (relativeCenter >= comfortStart && relativeCenter <= comfortEnd) return;
    const nextScrollTop = clampTreeScrollTop(
      host,
      itemCenter - host.clientHeight * OUTLINE_REVEAL_ZONE.target,
    );
    host.scrollTop = nextScrollTop;
    latestScrollTopRef.current = nextScrollTop;
    setScrollTop(nextScrollTop);
  }, [revealKey, rowIndexByKey, rows.length]);

  const ensureRowVisible = (rowIndex: number) => {
    const host = hostRef.current;
    if (!host) return;
    const rowTop =
      OUTLINE_TREE_TOP_PADDING + rowIndex * OUTLINE_TREE_ROW_HEIGHT;
    const rowBottom = rowTop + OUTLINE_TREE_ROW_HEIGHT;
    let nextScrollTop = host.scrollTop;
    if (rowTop < host.scrollTop) nextScrollTop = rowTop;
    else if (rowBottom > host.scrollTop + host.clientHeight) {
      nextScrollTop = rowBottom - host.clientHeight;
    }
    nextScrollTop = clampTreeScrollTop(host, nextScrollTop);
    if (nextScrollTop === host.scrollTop) return;
    host.scrollTop = nextScrollTop;
    latestScrollTopRef.current = nextScrollTop;
    setScrollTop(nextScrollTop);
  };

  const focusRow = (rowIndex: number) => {
    const row = rows[rowIndex];
    if (!row) return;
    pendingFocusKeyRef.current = row.key;
    setFocusedKey(row.key);
    ensureRowVisible(rowIndex);
    const renderedRow = rowRefs.current.get(row.key);
    if (renderedRow) {
      renderedRow.focus({ preventScroll: true });
      pendingFocusKeyRef.current = undefined;
    }
  };

  useLayoutEffect(() => {
    const pendingKey = pendingFocusKeyRef.current;
    if (!pendingKey) return;
    const row = rowRefs.current.get(pendingKey);
    if (!row) return;
    row.focus({ preventScroll: true });
    pendingFocusKeyRef.current = undefined;
  });

  const handleKeyDown = (
    event: KeyboardEvent<HTMLDivElement>,
    rowIndex: number,
  ) => {
    const row = rows[rowIndex];
    if (!row) return;
    const expandable = !row.node.isLeaf;
    const expanded = expandedKeySet.has(row.key);
    if (event.key === 'ArrowDown') focusRow(rowIndex + 1);
    else if (event.key === 'ArrowUp') focusRow(rowIndex - 1);
    else if (event.key === 'Home') focusRow(0);
    else if (event.key === 'End') focusRow(rows.length - 1);
    else if (event.key === 'ArrowRight') {
      if (expandable && !expanded) onToggle(row.key, true);
      else if (expanded && rows[rowIndex + 1]?.parentKey === row.key) {
        focusRow(rowIndex + 1);
      } else return;
    } else if (event.key === 'ArrowLeft') {
      if (expandable && expanded) onToggle(row.key, false);
      else if (row.parentKey) {
        const parentIndex = rowIndexByKey.get(row.parentKey);
        if (parentIndex !== undefined) focusRow(parentIndex);
      } else return;
    } else if (event.key === 'Enter' || event.key === ' ') {
      onSelect(row.key);
    } else return;
    event.preventDefault();
  };

  const handleScroll = (event: UIEvent<HTMLDivElement>) => {
    latestScrollTopRef.current = event.currentTarget.scrollTop;
    if (!virtual) return;
    if (typeof requestAnimationFrame === 'undefined') {
      setScrollTop(latestScrollTopRef.current);
      return;
    }
    if (scrollFrameRef.current !== undefined) return;
    scrollFrameRef.current = requestAnimationFrame(() => {
      scrollFrameRef.current = undefined;
      setScrollTop(latestScrollTopRef.current);
    });
  };

  const normalizedScrollTop = Math.max(0, scrollTop - OUTLINE_TREE_TOP_PADDING);
  const startIndex = virtual
    ? Math.max(
        0,
        Math.floor(normalizedScrollTop / OUTLINE_TREE_ROW_HEIGHT) -
          OUTLINE_TREE_OVERSCAN,
      )
    : 0;
  const endIndex = virtual
    ? Math.min(
        rows.length,
        Math.ceil(
          (normalizedScrollTop + viewportHeight) / OUTLINE_TREE_ROW_HEIGHT,
        ) + OUTLINE_TREE_OVERSCAN,
      )
    : rows.length;
  const renderedRows = rows.slice(startIndex, endIndex);

  const renderRow = (row: WordOutlineTreeRow, offset: number) => {
    const rowIndex = startIndex + offset;
    return (
      <OutlineTreeRowView
        key={row.key}
        row={row}
        rowIndex={rowIndex}
        expanded={expandedKeySet.has(row.key)}
        selected={selectedKey === row.key}
        focused={focusedKey === row.key}
        loading={loadingKeys.has(row.key)}
        expandLabel={expandLabel}
        collapseLabel={collapseLabel}
        registerRow={(key, element) => {
          if (element) rowRefs.current.set(key, element);
          else rowRefs.current.delete(key);
        }}
        onToggle={(currentRow) =>
          onToggle(currentRow.key, !expandedKeySet.has(currentRow.key))
        }
        onSelect={(currentRow, element) => {
          setFocusedKey(currentRow.key);
          element.focus({ preventScroll: true });
          onSelect(currentRow.key);
        }}
        onFocus={setFocusedKey}
        onKeyDown={handleKeyDown}
      />
    );
  };

  return (
    <div
      ref={hostRef}
      className="office-file-word-outline__tree-host"
      data-virtual={virtual ? 'true' : 'false'}
      onScroll={handleScroll}
    >
      {virtual ? (
        <div
          className="office-file-outline-tree office-file-outline-tree--virtual"
          style={{ height: rows.length * OUTLINE_TREE_ROW_HEIGHT }}
          aria-busy={loadingKeys.size > 0}
          aria-label={label}
          role="tree"
        >
          <div
            className="office-file-outline-tree__window"
            style={{
              transform: `translateY(${
                startIndex * OUTLINE_TREE_ROW_HEIGHT
              }px)`,
            }}
          >
            {renderedRows.map(renderRow)}
          </div>
        </div>
      ) : (
        <div
          className="office-file-outline-tree"
          aria-busy={loadingKeys.size > 0}
          aria-label={label}
          role="tree"
        >
          {renderedRows.map(renderRow)}
        </div>
      )}
    </div>
  );
}

export const OfficeOutlineTree = memo(OfficeOutlineTreeComponent);
