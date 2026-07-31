// XlsxSheetTabs 渲染工作表标签栏，并展示当前工作表范围或行列数量。
import { Tabs, Typography } from 'antd';
import React, { memo, useMemo } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type {
  SpreadsheetSheetDescriptor,
  SpreadsheetSourceSnapshot,
} from '../../services/spreadsheet/SpreadsheetSource';

/** Excel工作表标签栏组件属性。 */
type XlsxSheetTabsProps = {
  /** 当前数据源的只读快照。 */
  snapshot: SpreadsheetSourceSnapshot;
  /** 当前选中的工作表描述符。 */
  activeSheet: SpreadsheetSheetDescriptor;
  /** 在 SelectSheet 事件发生时调用的回调函数。 */
  onSelectSheet: (sheetId: string) => void;
};

/** 工作簿尚未加载时复用的空工作表页签集合。 */
const EMPTY_TABS: Array<{
  /** 用于稳定识别工作表标签项的键。 */
  key: string;
  /** 工作表标签显示的文本。 */
  label: string;
}> = [];

/** 渲染工作簿的工作表标签栏。 */
function XlsxSheetTabsComponent({
  snapshot,
  activeSheet,
  onSelectSheet,
}: XlsxSheetTabsProps) {
  const messages = useOfficeFileViewerMessages();
  const tabItems = useMemo(
    () =>
      snapshot.sheets.map((sheet) => ({
        key: sheet.id,
        label: sheet.name,
      })) ?? EMPTY_TABS,
    [snapshot],
  );
  const rangeText = messages.spreadsheet.dimensions(
    activeSheet.rowCount,
    activeSheet.columnCount,
  );

  return (
    <div className="office-file-xlsx-sheet-tabs">
      <Tabs
        activeKey={activeSheet.id}
        onChange={onSelectSheet}
        items={tabItems}
        tabBarExtraContent={
          <Typography.Text
            type="secondary"
            className="office-file-xlsx-sheet-tabs__range"
          >
            {rangeText}
          </Typography.Text>
        }
      />
    </div>
  );
}

export const XlsxSheetTabs = memo(XlsxSheetTabsComponent);
