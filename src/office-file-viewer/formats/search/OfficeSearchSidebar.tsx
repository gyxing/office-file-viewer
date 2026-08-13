import type { KeyboardEvent } from 'react';
import React, { memo, useEffect, useRef } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import {
  ChevronDownIcon,
  ChevronUpIcon,
  CloseIcon,
  SearchIcon,
} from '../../shared/ui/OfficeIcons';
import { OfficeNavigationModeTabs } from '../navigation/OfficeNavigationModeTabs';
import { OfficeNavigationPanel } from '../navigation/OfficeNavigationPanel';
import './index.less';
import { OfficeSearchResultList } from './OfficeSearchResultList';
import type { OfficeSearchController } from './useOfficeSearchController';

/** 文档查找侧栏属性。 */
type OfficeSearchSidebarProps = {
  /** 当前是否显示查找侧栏。 */
  visible: boolean;
  /** 当前解析会话的稳定标识。 */
  sessionKey?: string;
  /** 当前查看器独占的搜索控制器。 */
  controller: OfficeSearchController;
  /** 关闭查找侧栏。 */
  onClose(): void;
  /** Word 文档存在大纲时切换到大纲侧栏。 */
  onShowOutline?: () => void;
};

/** 渲染搜索输入、选项、结果导航和按需虚拟化结果列表。 */
function OfficeSearchSidebarComponent({
  visible,
  sessionKey,
  controller,
  onClose,
  onShowOutline,
}: OfficeSearchSidebarProps) {
  const messages = useOfficeFileViewerMessages();
  const inputRef = useRef<HTMLInputElement>(null);
  const { state, actions } = controller;

  useEffect(() => {
    if (!visible) return;
    const frame = requestAnimationFrame(() => {
      inputRef.current?.focus({ preventScroll: true });
      inputRef.current?.select();
    });
    return () => cancelAnimationFrame(frame);
  }, [visible]);

  const handleInputKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      void (event.shiftKey ? actions.previousResult() : actions.nextResult());
    } else if (event.key === 'Escape') {
      event.preventDefault();
      onClose();
    }
  };
  const currentPosition =
    state.results.length && state.currentIndex >= 0
      ? messages.search.currentResult(
          state.currentIndex + 1,
          state.results.length,
        )
      : messages.search.resultCount(state.results.length);
  const status = state.error
    ? state.error
    : !state.query
    ? messages.search.emptyQuery
    : state.searching
    ? messages.search.searching
    : state.results.length
    ? messages.search.resultCount(state.results.length)
    : messages.search.noResults;

  return (
    <OfficeNavigationPanel
      visible={visible}
      sessionKey={sessionKey ?? 'empty-search-session'}
      ariaLabel={messages.search.region}
      resizeLabel={messages.search.resize}
      className="office-file-search-sidebar"
      dataAttributes={{
        'data-testid': 'office-search-sidebar',
        'data-search-complete': state.complete ? 'true' : 'false',
        'data-result-count': state.results.length,
        'data-current-index': state.currentIndex,
      }}
    >
      <div className="office-file-search-sidebar__surface">
        <header className="office-file-search-sidebar__header">
          {onShowOutline ? (
            <OfficeNavigationModeTabs
              activeMode="search"
              outlineLabel={messages.outline.title}
              searchLabel={messages.search.title}
              onShowOutline={onShowOutline}
              onShowSearch={() => undefined}
            />
          ) : (
            <span className="office-file-search-sidebar__title">
              <SearchIcon />
              {messages.search.title}
            </span>
          )}
          <span className="office-file-search-sidebar__position">
            {currentPosition}
          </span>
          <button
            type="button"
            className="office-file-search-sidebar__icon-button"
            aria-label={messages.search.collapse}
            title={messages.search.collapse}
            onClick={onClose}
          >
            <CloseIcon />
          </button>
        </header>
        <div className="office-file-search-sidebar__query-row">
          <div className="office-file-search-sidebar__input-wrap">
            <SearchIcon />
            <input
              ref={inputRef}
              type="search"
              value={state.query}
              aria-label={messages.search.placeholder}
              placeholder={messages.search.placeholder}
              onChange={(event) => actions.setQuery(event.currentTarget.value)}
              onKeyDown={handleInputKeyDown}
            />
          </div>
          <button
            type="button"
            className="office-file-search-sidebar__icon-button"
            aria-label={messages.search.previous}
            title={messages.search.previous}
            disabled={!state.results.length}
            onClick={() => void actions.previousResult()}
          >
            <ChevronUpIcon />
          </button>
          <button
            type="button"
            className="office-file-search-sidebar__icon-button"
            aria-label={messages.search.next}
            title={messages.search.next}
            disabled={!state.results.length}
            onClick={() => void actions.nextResult()}
          >
            <ChevronDownIcon />
          </button>
        </div>
        <div className="office-file-search-sidebar__options">
          <label>
            <input
              type="checkbox"
              checked={state.matchCase}
              onChange={(event) =>
                actions.setMatchCase(event.currentTarget.checked)
              }
            />
            {messages.search.matchCase}
          </label>
          <label>
            <input
              type="checkbox"
              checked={state.wholeWord}
              onChange={(event) =>
                actions.setWholeWord(event.currentTarget.checked)
              }
            />
            {messages.search.wholeWord}
          </label>
        </div>
        <div
          className="office-file-search-sidebar__status"
          role="status"
          aria-live="polite"
        >
          {status}
        </div>
        {state.results.length ? (
          <OfficeSearchResultList
            results={state.results}
            currentIndex={state.currentIndex}
            onSelect={(index) => void actions.selectResult(index)}
          />
        ) : (
          <div className="office-file-search-sidebar__empty">{status}</div>
        )}
      </div>
    </OfficeNavigationPanel>
  );
}

export const OfficeSearchSidebar = memo(OfficeSearchSidebarComponent);
