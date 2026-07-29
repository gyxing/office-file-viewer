import React, { memo } from 'react';
import type { DocxChartBlock } from '../../services/docx/types';
import { useOfficeResourceUrl } from '../../services/resource-store/useOfficeResourceUrl';
import { OfficeChartView } from '../../shared/chart/OfficeChartView';

type DocxChartViewProps = {
  block: DocxChartBlock;
  zoom: number;
};

/** 为 DOCX 图表统一获取按页快照资源，同时保留 ECharts 的卸载生命周期。 */
function DocxChartViewComponent({ block, zoom }: DocxChartViewProps) {
  const snapshot = useOfficeResourceUrl(block.snapshotSource);
  return (
    <OfficeChartView
      chart={
        snapshot.url
          ? { ...block.chart, snapshotSrc: snapshot.url }
          : block.chart
      }
      width={block.width}
      height={block.height}
      zoom={zoom}
    />
  );
}

export const DocxChartView = memo(DocxChartViewComponent);
