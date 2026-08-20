import type { KeyboardEvent, UIEvent } from 'react';
import React, { memo, useEffect, useMemo, useRef, useState } from 'react';
import { useOfficeNavigationResize } from '../../formats/navigation/useOfficeNavigationResize';
import { useOfficeFileViewerMessages } from '../../locale';
import type { OfficeAnnotation } from '../../services/annotations/types';
import { OfficeButton } from '../ui/OfficeButton';
import { ChevronDownIcon, ChevronUpIcon, CloseIcon } from '../ui/OfficeIcons';
import './index.less';
import type { OfficeAnnotationController } from './useOfficeAnnotationController';

/** 虚拟列表使用的批注卡片固定占位高度。 */
const ANNOTATION_ITEM_HEIGHT = 112;
/** 视口上下额外保留的批注数量，避免快速滚动出现空白。 */
const ANNOTATION_LIST_OVERSCAN = 5;
/** 右侧审阅面板使用反向拖动和独立宽度变量。 */
const REVIEW_RESIZE_OPTIONS = {
  growDirection: 'left',
  widthCssVariable: '--office-file-review-panel-width',
} as const;

/** 审阅面板属性。 */
type OfficeAnnotationPanelProps = {
  /** 当前是否显示审阅面板。 */
  visible: boolean;
  /** 切换文档时用于清空列表和恢复宽度的稳定标识。 */
  sessionKey: string;
  /** 当前查看器独占的批注控制器。 */
  controller: OfficeAnnotationController;
  /** 关闭审阅面板。 */
  onClose(): void;
};

/** 为不同格式生成简短且稳定的批注目标说明。 */
function getAnnotationTargetLabel(
  annotation: OfficeAnnotation,
  messages: ReturnType<typeof useOfficeFileViewerMessages>,
) {
  switch (annotation.target.kind) {
    case 'word-range':
      return annotation.target.pageIndex === undefined
        ? messages.review.wordTarget
        : messages.review.pageTarget(annotation.target.pageIndex + 1);
    case 'spreadsheet-cell':
      return messages.review.cellTarget(
        annotation.target.sheetId,
        annotation.target.row,
        annotation.target.column,
      );
    case 'presentation-element':
      return annotation.target.slideIndex === undefined
        ? messages.review.slideTarget(annotation.target.slideId)
        : messages.review.pageTarget(annotation.target.slideIndex + 1);
  }
}

/** 渲染单个固定高度批注卡片，长正文通过标题保留完整可读内容。 */
function AnnotationListItem({
  annotation,
  index,
  active,
  onSelect,
}: {
  /** 当前索引已经加载的批注。 */
  annotation: OfficeAnnotation;
  /** 当前批注的零基索引。 */
  index: number;
  /** 当前批注是否已经激活。 */
  active: boolean;
  /** 请求选中当前批注。 */
  onSelect(index: number): void;
}) {
  const messages = useOfficeFileViewerMessages();
  return (
    <button
      className="office-file-annotation-panel__item"
      type="button"
      role="option"
      aria-selected={active}
      data-annotation-index={index}
      onClick={() => onSelect(index)}
    >
      <span className="office-file-annotation-panel__item-meta">
        <strong>{annotation.author || messages.review.unknownAuthor}</strong>
        <span>{getAnnotationTargetLabel(annotation, messages)}</span>
      </span>
      <span
        className="office-file-annotation-panel__item-text"
        title={annotation.text}
      >
        {annotation.text || messages.review.emptyComment}
      </span>
      <span className="office-file-annotation-panel__item-footer">
        <span>{annotation.createdAt}</span>
        {annotation.resolved ? (
          <span className="office-file-annotation-panel__resolved">
            {messages.review.resolved}
          </span>
        ) : null}
      </span>
    </button>
  );
}

/** 只加载并渲染当前视口附近的批注卡片。 */
function VirtualAnnotationList({
  controller,
}: {
  /** 当前查看器独占的批注控制器。 */
  controller: OfficeAnnotationController;
}) {
  const messages = useOfficeFileViewerMessages();
  const viewportRef = useRef<HTMLDivElement>(null);
  const [scrollTop, setScrollTop] = useState(0);
  const [viewportHeight, setViewportHeight] = useState(360);
  const { snapshot, activeIndex, cacheRevision, loading } = controller.state;
  const count = snapshot.count;
  const windowRange = useMemo(() => {
    const start = Math.max(
      0,
      Math.floor(scrollTop / ANNOTATION_ITEM_HEIGHT) - ANNOTATION_LIST_OVERSCAN,
    );
    const end = Math.min(
      count,
      Math.ceil((scrollTop + viewportHeight) / ANNOTATION_ITEM_HEIGHT) +
        ANNOTATION_LIST_OVERSCAN,
    );
    return { start, end };
  }, [count, scrollTop, viewportHeight]);

  useEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return undefined;
    const updateHeight = () => setViewportHeight(viewport.clientHeight || 360);
    updateHeight();
    if (typeof ResizeObserver === 'undefined') return undefined;
    const observer = new ResizeObserver(updateHeight);
    observer.observe(viewport);
    return () => observer.disconnect();
  }, []);

  useEffect(() => {
    void controller.actions.loadRange(windowRange.start, windowRange.end);
  }, [
    controller.actions,
    snapshot.revision,
    windowRange.end,
    windowRange.start,
  ]);

  useEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport || activeIndex < 0) return;
    const targetTop = activeIndex * ANNOTATION_ITEM_HEIGHT;
    const nextTop = Math.max(
      0,
      Math.min(
        count * ANNOTATION_ITEM_HEIGHT - viewport.clientHeight,
        targetTop - (viewport.clientHeight - ANNOTATION_ITEM_HEIGHT) / 2,
      ),
    );
    viewport.scrollTo({ top: nextTop, behavior: 'smooth' });
  }, [activeIndex, count]);

  const handleScroll = (event: UIEvent<HTMLDivElement>) => {
    setScrollTop(event.currentTarget.scrollTop);
  };
  const handleKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'ArrowUp') {
      event.preventDefault();
      void controller.actions.previous();
    } else if (event.key === 'ArrowDown') {
      event.preventDefault();
      void controller.actions.next();
    } else if (event.key === 'Home') {
      event.preventDefault();
      void controller.actions.selectIndex(0);
    } else if (event.key === 'End') {
      event.preventDefault();
      void controller.actions.selectIndex(count - 1);
    }
  };

  const items = [];
  for (let index = windowRange.start; index < windowRange.end; index += 1) {
    const annotation = controller.actions.getCached(index);
    if (!annotation) continue;
    items.push(
      <div
        key={annotation.id}
        className="office-file-annotation-panel__item-position"
        style={{ transform: `translateY(${index * ANNOTATION_ITEM_HEIGHT}px)` }}
      >
        <AnnotationListItem
          annotation={annotation}
          index={index}
          active={activeIndex === index}
          onSelect={(nextIndex) =>
            void controller.actions.selectIndex(nextIndex)
          }
        />
      </div>,
    );
  }

  return (
    <div
      ref={viewportRef}
      className="office-file-annotation-panel__list"
      role="listbox"
      aria-label={messages.review.comments}
      aria-busy={loading}
      tabIndex={0}
      onScroll={handleScroll}
      onKeyDown={handleKeyDown}
      data-cache-revision={cacheRevision}
    >
      <div
        className="office-file-annotation-panel__list-space"
        style={{ height: count * ANNOTATION_ITEM_HEIGHT }}
      >
        {items}
      </div>
    </div>
  );
}

/** 渲染可动画显隐、可调整宽度的右侧审阅面板。 */
function OfficeAnnotationPanelComponent({
  visible,
  sessionKey,
  controller,
  onClose,
}: OfficeAnnotationPanelProps) {
  const messages = useOfficeFileViewerMessages();
  const panelRef = useRef<HTMLElement>(null);
  const { snapshot } = controller.state;
  const commentCount = controller.options.showComments ? snapshot.count : 0;
  const noteCount = controller.options.showNotes ? snapshot.noteCount : 0;
  const resize = useOfficeNavigationResize(
    panelRef,
    sessionKey,
    REVIEW_RESIZE_OPTIONS,
  );

  useEffect(() => {
    const panel = panelRef.current;
    if (!panel) return;
    if (visible) panel.removeAttribute('inert');
    else panel.setAttribute('inert', '');
  }, [visible]);

  const hasSummary = noteCount > 0;
  return (
    <aside
      ref={panelRef}
      className="office-file-annotation-panel"
      aria-label={messages.review.region}
      aria-hidden={!visible}
      data-visible={visible ? 'true' : 'false'}
      data-annotation-count={commentCount}
      data-revision-count={snapshot.revisionCount}
    >
      <div
        ref={resize.handleRef}
        className="office-file-annotation-panel__resize-handle"
        role="separator"
        aria-label={messages.review.resize}
        aria-orientation="vertical"
        aria-valuemin={resize.minWidth}
        aria-valuemax={Math.round(resize.maxWidth)}
        aria-valuenow={Math.round(resize.width)}
        title={messages.review.resize}
        tabIndex={0}
        onKeyDown={resize.handleKeyDown}
        onPointerDown={resize.handlePointerDown}
        onPointerMove={resize.handlePointerMove}
        onPointerUp={resize.handlePointerEnd}
        onPointerCancel={resize.handlePointerEnd}
        onLostPointerCapture={resize.handlePointerEnd}
      >
        <span aria-hidden="true" />
      </div>
      <div className="office-file-annotation-panel__surface">
        <header className="office-file-annotation-panel__header">
          <div>
            <strong>{messages.review.title}</strong>
            <span>{messages.review.itemCount(commentCount)}</span>
          </div>
          <div className="office-file-annotation-panel__header-actions">
            <OfficeButton
              size="small"
              aria-label={messages.review.previous}
              icon={<ChevronUpIcon />}
              disabled={!commentCount}
              onClick={() => void controller.actions.previous()}
            />
            <OfficeButton
              size="small"
              aria-label={messages.review.next}
              icon={<ChevronDownIcon />}
              disabled={!commentCount}
              onClick={() => void controller.actions.next()}
            />
            <OfficeButton
              size="small"
              aria-label={messages.review.close}
              icon={<CloseIcon />}
              onClick={onClose}
            />
          </div>
        </header>
        {hasSummary ? (
          <div className="office-file-annotation-panel__summary">
            {noteCount ? (
              <span>{messages.review.noteCount(noteCount)}</span>
            ) : null}
          </div>
        ) : null}
        {controller.state.error ? (
          <div className="office-file-annotation-panel__status" role="alert">
            {controller.state.error}
          </div>
        ) : commentCount ? (
          <VirtualAnnotationList controller={controller} />
        ) : (
          <div className="office-file-annotation-panel__status">
            {messages.review.empty}
          </div>
        )}
      </div>
    </aside>
  );
}

export const OfficeAnnotationPanel = memo(OfficeAnnotationPanelComponent);
