// XlsxSheetTabs 渲染工作表标签栏，并展示当前工作表范围或行列数量。
import { Tabs, Typography } from 'antd';
import React, { memo, useMemo } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type {
  SpreadsheetSheetDescriptor,
  SpreadsheetSourceSnapshot,
} from '../../services/spreadsheet/SpreadsheetSource';

/** 定义 XlsxSheetTabs 组件可接收的属性。 */
type XlsxSheetTabsProps = {
  /** XlsxSheetTabsProps 当前关联的标准化工作簿。 */
  snapshot: SpreadsheetSourceSnapshot;
  /** 当前选中的工作表描述符。 */
  activeSheet: SpreadsheetSheetDescriptor;
  /** 在 SelectSheet 事件发生时调用的回调函数。 */
  onSelectSheet: (sheetId: string) => void;
};

const EMPTY_TABS: Array<{
  /** 当前内联结构 在界面列表或映射中的稳定键。 */ key: string;
  /** 当前内联结构 面向用户展示的标签文本。 */
  label: string;
}> = [];

/** 渲染 XlsxSheetTabsComponent 组件。 */
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
