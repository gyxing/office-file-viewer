// XlsxViewer 负责 XLSX 工作簿预览的整体布局，包括工作表选择和当前工作表内容区。
import React, { memo, useMemo } from 'react';
import type { SpreadsheetSource } from '../../services/spreadsheet/SpreadsheetSource';
import { getSpreadsheetSource } from '../../services/spreadsheet/spreadsheetSourceRegistry';
import type { XlsxWorkbook } from '../../services/xlsx/types';
import { OfficeEmpty } from '../../shell/Empty';
import './index.less';
import { SpreadsheetSheetState } from './SpreadsheetSheetState';
import { useSpreadsheetSource } from './useSpreadsheetSource';
import { VirtualSpreadsheetGrid } from './VirtualSpreadsheetGrid';
import { XlsxChartSheet } from './XlsxChartSheet';
import { XlsxSheetGrid } from './XlsxSheetGrid';
import { XlsxSheetTabs } from './XlsxSheetTabs';

/** Excel预览器组件属性。 */
type XlsxViewerProps = {
  /** 当前处理的标准化工作簿。 */
  workbook?: XlsxWorkbook;
  /** 大型 XLS/XLSX 使用的按 Sheet 数据源。 */
  source?: SpreadsheetSource;
  /** 当前模型对应的 Office 内容类型。 */
  kind?: 'xlsx' | 'xls';
  /** 当前激活工作表的稳定标识。 */
  activeSheetId?: string;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 在 SelectSheet 事件发生时调用的回调函数。 */
  onSelectSheet: (sheetId: string) => void;
};

/** 渲染Excel预览器。 */
function XlsxViewerComponent({
  workbook,
  source,
  kind = 'xlsx',
  activeSheetId,
  zoom,
  onSelectSheet,
}: XlsxViewerProps) {
  const resolvedSource = useMemo(
    () => source ?? (workbook ? getSpreadsheetSource(workbook) : undefined),
    [source, workbook],
  );

  if (!resolvedSource) {
    return <OfficeEmpty kind={kind} />;
  }

  return (
    <XlsxSourceViewer
      source={resolvedSource}
      kind={kind}
      activeSheetId={activeSheetId}
      zoom={zoom}
      onSelectSheet={onSelectSheet}
    />
  );
}

/** 渲染已经统一为 Source 的工作簿。 */
function XlsxSourceViewer({
  source,
  kind,
  activeSheetId,
  zoom,
  onSelectSheet,
}: {
  source: SpreadsheetSource;
  kind: 'xlsx' | 'xls';
  activeSheetId?: string;
  zoom: number;
  onSelectSheet: (sheetId: string) => void;
}) {
  const state = useSpreadsheetSource(source, activeSheetId);
  const descriptor = state.activeDescriptor;

  if (!descriptor) return <OfficeEmpty kind={kind} />;

  return (
    <div className="office-file-xlsx-viewer">
      <XlsxSheetTabs
        snapshot={state.snapshot}
        activeSheet={descriptor}
        onSelectSheet={onSelectSheet}
      />
      <SpreadsheetSheetState
        loading={state.loading}
        error={state.error}
        retry={state.retry}
      >
        {state.activeSheet?.kind === 'chart' ? (
          <XlsxChartSheet sheet={state.activeSheet} zoom={zoom} />
        ) : state.profile?.gridMode === 'table' && state.activeSheet ? (
          <XlsxSheetGrid sheet={state.activeSheet} zoom={zoom} />
        ) : state.profile ? (
          <VirtualSpreadsheetGrid
            source={source}
            sheetId={descriptor.id}
            layout={source.getSheetLayout(descriptor.id)}
            gridMode={state.profile.gridMode}
            zoom={zoom}
          />
        ) : null}
      </SpreadsheetSheetState>
    </div>
  );
}

export const XlsxViewer = memo(XlsxViewerComponent);
