import type { KeyboardEvent } from 'react';
import React, { memo } from 'react';

/** Word 左侧导航在大纲与查找之间切换时使用的页签属性。 */
type OfficeNavigationModeTabsProps = {
  /** 当前显示的导航模式。 */
  activeMode: 'outline' | 'search';
  /** 大纲页签文案。 */
  outlineLabel: string;
  /** 查找页签文案。 */
  searchLabel: string;
  /** 请求显示大纲。 */
  onShowOutline(): void;
  /** 请求显示查找。 */
  onShowSearch(): void;
};

/** 渲染 Word 大纲与查找的互斥模式页签。 */
function OfficeNavigationModeTabsComponent({
  activeMode,
  outlineLabel,
  searchLabel,
  onShowOutline,
  onShowSearch,
}: OfficeNavigationModeTabsProps) {
  const handleKeyDown = (event: KeyboardEvent<HTMLButtonElement>) => {
    let nextMode: OfficeNavigationModeTabsProps['activeMode'] | undefined;
    if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
      nextMode = activeMode === 'outline' ? 'search' : 'outline';
    } else if (event.key === 'Home') {
      nextMode = 'outline';
    } else if (event.key === 'End') {
      nextMode = 'search';
    }
    if (!nextMode) return;
    event.preventDefault();
    if (nextMode === activeMode) return;
    if (nextMode === 'outline') onShowOutline();
    else onShowSearch();
  };

  return (
    <div
      className="office-file-navigation-tabs"
      role="tablist"
      aria-label={`${outlineLabel} / ${searchLabel}`}
    >
      <button
        type="button"
        className="office-file-navigation-tabs__tab"
        role="tab"
        aria-selected={activeMode === 'outline'}
        tabIndex={activeMode === 'outline' ? 0 : -1}
        onClick={onShowOutline}
        onKeyDown={handleKeyDown}
      >
        {outlineLabel}
      </button>
      <button
        type="button"
        className="office-file-navigation-tabs__tab"
        role="tab"
        aria-selected={activeMode === 'search'}
        tabIndex={activeMode === 'search' ? 0 : -1}
        onClick={onShowSearch}
        onKeyDown={handleKeyDown}
      >
        {searchLabel}
      </button>
    </div>
  );
}

export const OfficeNavigationModeTabs = memo(OfficeNavigationModeTabsComponent);
