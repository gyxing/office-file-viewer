import type { KeyboardEvent } from 'react';
import React, { memo, useEffect, useMemo, useRef } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type {
  SpreadsheetSheetDescriptor,
  SpreadsheetSourceSnapshot,
} from '../../services/spreadsheet/SpreadsheetSource';

/** Excel 工作表标签栏组件属性。 */
type XlsxSheetTabsProps = {
  /** 当前数据源的只读快照。 */
  snapshot: SpreadsheetSourceSnapshot;
  /** 当前选中的工作表描述符。 */
  activeSheet: SpreadsheetSheetDescriptor;
  /** 用户选择工作表时调用的回调函数。 */
  onSelectSheet: (sheetId: string) => void;
};

/** 渲染支持横向滚动和键盘导航的工作表标签栏。 */
function XlsxSheetTabsComponent({
  snapshot,
  activeSheet,
  onSelectSheet,
}: XlsxSheetTabsProps) {
  const messages = useOfficeFileViewerMessages();
  const listRef = useRef<HTMLDivElement>(null);
  const tabRefs = useRef(new Map<string, HTMLButtonElement>());
  const tabItems = useMemo(
    () =>
      snapshot.sheets.map((sheet) => ({
        key: sheet.id,
        label: sheet.name,
      })),
    [snapshot.sheets],
  );
  const activeIndex = tabItems.findIndex((item) => item.key === activeSheet.id);
  const rangeText = messages.spreadsheet.dimensions(
    activeSheet.rowCount,
    activeSheet.columnCount,
  );

  useEffect(() => {
    const listElement = listRef.current;
    const activeTabElement = tabRefs.current.get(activeSheet.id);
    if (!listElement || !activeTabElement) return;

    // 仅调整标签栏自身的横向滚动，避免 scrollIntoView 连带移动预览页面。
    const tabStart = activeTabElement.offsetLeft;
    const tabEnd = tabStart + activeTabElement.offsetWidth;
    const visibleStart = listElement.scrollLeft;
    const visibleEnd = visibleStart + listElement.clientWidth;
    if (tabStart < visibleStart) listElement.scrollLeft = tabStart;
    else if (tabEnd > visibleEnd) {
      listElement.scrollLeft = tabEnd - listElement.clientWidth;
    }
  }, [activeSheet.id]);

  const selectAndFocus = (index: number) => {
    const item = tabItems[index];
    if (!item) return;
    onSelectSheet(item.key);
    tabRefs.current.get(item.key)?.focus();
  };

  const handleKeyDown = (
    event: KeyboardEvent<HTMLButtonElement>,
    index: number,
  ) => {
    let nextIndex: number | undefined;
    if (event.key === 'ArrowLeft') {
      nextIndex = index > 0 ? index - 1 : tabItems.length - 1;
    } else if (event.key === 'ArrowRight') {
      nextIndex = index < tabItems.length - 1 ? index + 1 : 0;
    } else if (event.key === 'Home') {
      nextIndex = 0;
    } else if (event.key === 'End') {
      nextIndex = tabItems.length - 1;
    }

    if (nextIndex === undefined) return;
    event.preventDefault();
    selectAndFocus(nextIndex);
  };

  return (
    <div className="office-file-xlsx-sheet-tabs">
      <div
        ref={listRef}
        className="office-file-xlsx-sheet-tabs__list"
        aria-label={messages.spreadsheet.sheets}
        role="tablist"
      >
        {tabItems.map((item, index) => {
          const active = index === activeIndex;
          return (
            <button
              key={item.key}
              ref={(element) => {
                if (element) tabRefs.current.set(item.key, element);
                else tabRefs.current.delete(item.key);
              }}
              className="office-file-xlsx-sheet-tabs__tab"
              type="button"
              aria-selected={active}
              role="tab"
              tabIndex={active ? 0 : -1}
              title={item.label}
              onClick={() => onSelectSheet(item.key)}
              onKeyDown={(event) => handleKeyDown(event, index)}
            >
              {item.label}
            </button>
          );
        })}
      </div>
      <span className="office-file-xlsx-sheet-tabs__range">{rangeText}</span>
    </div>
  );
}

export const XlsxSheetTabs = memo(XlsxSheetTabsComponent);
