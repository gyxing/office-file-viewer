// XlsxViewer 负责 XLSX 工作簿预览的整体布局，包括工作表选择和当前工作表内容区。
import React, { memo, useMemo } from 'react';
import type { XlsxWorkbook } from '../../services/xlsx/types';
import { OfficeEmpty } from '../../shell/Empty';
import './index.less';
import { XlsxChartSheet } from './XlsxChartSheet';
import { XlsxSheetGrid } from './XlsxSheetGrid';
import { XlsxSheetTabs } from './XlsxSheetTabs';

/** 定义 XlsxViewer 组件可接收的属性。 */
type XlsxViewerProps = {
  /** XlsxViewerProps 当前关联的标准化工作簿。 */
  workbook?: XlsxWorkbook;
  /** 标识 XlsxViewerProps 对应的 Office 文件或数据种类。 */
  kind?: 'xlsx' | 'xls';
  /** XlsxViewerProps 的 activeSheetId 文本值。 */
  activeSheetId?: string;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 在 SelectSheet 事件发生时调用的回调函数。 */
  onSelectSheet: (sheetId: string) => void;
};

/** 渲染 XlsxViewerComponent 组件。 */
function XlsxViewerComponent({
  workbook,
  kind = 'xlsx',
  activeSheetId,
  zoom,
  onSelectSheet,
}: XlsxViewerProps) {
  const activeSheet = useMemo(
    () =>
      workbook?.sheets.find((sheet) => sheet.id === activeSheetId) ??
      workbook?.sheets[0],
    [activeSheetId, workbook],
  );

  if (!activeSheet) {
    return <OfficeEmpty kind={kind} />;
  }

  return (
    <div className="office-file-xlsx-viewer">
      <XlsxSheetTabs
        workbook={workbook}
        activeSheet={activeSheet}
        onSelectSheet={onSelectSheet}
      />
      {activeSheet.kind === 'chart' ? (
        <XlsxChartSheet sheet={activeSheet} zoom={zoom} />
      ) : (
        <XlsxSheetGrid sheet={activeSheet} zoom={zoom} />
      )}
    </div>
  );
}

export const XlsxViewer = memo(XlsxViewerComponent);
