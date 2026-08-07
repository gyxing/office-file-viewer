import type { MutableRefObject, RefObject } from 'react';
import React, { memo, useEffect, useMemo, useRef, useState } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type {
  WordOutlineProvider,
  WordOutlineProviderSnapshot,
} from '../../services/word/WordOutlineProvider';
import type { WordPageSource } from '../../services/word/WordPageSource';
import type { WordOutlineItem } from '../../services/word/types';
import { useExternalStoreSnapshot } from '../../shared/react/useExternalStoreSnapshot';
import { OutlineIcon, PanelLeftCloseIcon } from '../../shared/ui/OfficeIcons';
import type { WordBlockPageIndex } from '../word-pages/WordBlockPageIndex';
import type { WordPageNavigationController } from '../word-pages/types';
import { WordOutlineTree } from './WordOutlineTree';
import './index.less';
import { collectOutlineTreeKeys } from './outlineTreeModel';
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
  /** 支持普通完整树和大文件按层读取的大纲数据源。 */
  provider: WordOutlineProvider;
  /** 当前画像选择的大纲渲染路径。 */
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

/** 未首次展开时不订阅渐进式大纲，避免隐藏侧栏参与解析期重渲染。 */
const INACTIVE_OUTLINE_SNAPSHOT: WordOutlineProviderSnapshot = {
  revision: 0,
  count: 0,
  complete: true,
};

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
        ? collectOutlineTreeKeys(normalTree)
        : provider.getRoots(1).map((node) => node.key),
    [normalTree, outlineMode, provider, snapshot.revision],
  );
  const [expandedKeys, setExpandedKeys] =
    useState<string[]>(initialExpandedKeys);
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

  const handleToggle = (key: string, expanded: boolean) => {
    setExpandedKeys((current) => {
      if (expanded) {
        return current.includes(key) ? current : [...current, key];
      }
      return current.includes(key)
        ? current.filter((currentKey) => currentKey !== key)
        : current;
    });
  };

  const handleSelect = (key: string) => {
    const selected = provider.locate(key);
    if (!selected) return;
    manuallySelectedKeyRef.current = selected.id;
    selectTarget(selected);
  };

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
            <WordOutlineTree
              normalTreeData={normalTree}
              provider={provider}
              virtual={outlineMode === 'virtual'}
              documentSessionId={documentSessionId}
              revision={snapshot.revision}
              expandedKeys={expandedKeys}
              selectedKey={activeKey}
              revealKey={revealKey}
              label={messages.outline.tree}
              expandLabel={messages.outline.expand}
              collapseLabel={messages.outline.collapse}
              onToggle={handleToggle}
              onSelect={handleSelect}
            />
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
