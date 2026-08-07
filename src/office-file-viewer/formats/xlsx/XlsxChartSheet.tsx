// XlsxChartSheet 让独立图表工作表占满内容区，不显示单元格表头。
import React, { memo } from 'react';
import type { XlsxSheet } from '../../services/xlsx/types';
import { OfficeChartView } from '../../shared/chart/OfficeChartView';
import { OfficePreviewEmpty } from '../common/OfficePreviewEmpty';

/** Excel图表工作表组件属性。 */
type XlsxChartSheetProps = {
  /** 当前处理的工作表。 */
  sheet: XlsxSheet;
  /** 当前预览缩放比例。 */
  zoom: number;
};

/** 渲染Excel图表工作表。 */
function XlsxChartSheetComponent({ sheet, zoom }: XlsxChartSheetProps) {
  const chart = sheet.charts[0];
  if (!chart) return <OfficePreviewEmpty kind="xls" />;
  return (
    <div className="office-file-xlsx-chart-sheet">
      <OfficeChartView
        chart={chart.chart}
        width={chart.width}
        height={chart.height}
        zoom={zoom}
      />
    </div>
  );
}

export const XlsxChartSheet = memo(XlsxChartSheetComponent);
