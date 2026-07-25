// XlsxSheetTabs 渲染工作表标签栏，并展示当前工作表范围或行列数量。
import { Tabs, Typography } from 'antd';
import React, { memo, useMemo } from 'react';
import type { XlsxSheet, XlsxWorkbook } from '../../services/xlsx/types';

/** 定义 XlsxSheetTabs 组件可接收的属性。 */
type XlsxSheetTabsProps = {
  /** XlsxSheetTabsProps 当前关联的标准化工作簿。 */
  workbook?: XlsxWorkbook;
  /** 当前选中的工作表模型。 */
  activeSheet: XlsxSheet;
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
  workbook,
  activeSheet,
  onSelectSheet,
}: XlsxSheetTabsProps) {
  const tabItems = useMemo(
    () =>
      workbook?.sheets.map((sheet) => ({
        key: sheet.id,
        label: sheet.name,
      })) ?? EMPTY_TABS,
    [workbook],
  );
  const rangeText =
    activeSheet.range ??
    `${activeSheet.rowCount} 行 x ${activeSheet.columnCount} 列`;

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
