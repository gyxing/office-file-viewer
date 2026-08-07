// XlsxSheetGrid 负责工作表滚动画布，统一承载表格、浮动图片和浮动图表。
import type { CSSProperties } from 'react';
import React, { memo, useEffect, useMemo, useRef, useState } from 'react';
import type { SpreadsheetViewMode } from '../../services/spreadsheet/viewMode';
import type { XlsxSheet } from '../../services/xlsx/types';
import { getXlsxSheetMetrics } from './sheetRenderUtils';
import { useXlsxSheetTableLayout } from './useXlsxSheetTableLayout';
import { XlsxFloatingCharts } from './XlsxFloatingCharts';
import { XlsxFloatingImages } from './XlsxFloatingImages';
import { XlsxSheetFiller } from './XlsxSheetFiller';
import { XlsxSheetTable } from './XlsxSheetTable';

/** Excel工作表网格组件属性。 */
type XlsxSheetGridProps = {
  /** 当前处理的工作表。 */
  sheet: XlsxSheet;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 当前电子表格采用的显示模式。 */
  viewMode: SpreadsheetViewMode;
};

/** 描述工作表滚动内容区当前可用的 CSS 像素尺寸。 */
type XlsxSheetViewportSize = {
  /** 滚动内容区不含内边距的宽度。 */
  width: number;
  /** 滚动内容区不含内边距的高度。 */
  height: number;
};

/** 电子表格网格尚未测量时使用的零尺寸。 */
const EMPTY_VIEWPORT_SIZE: XlsxSheetViewportSize = { width: 0, height: 0 };

/** 必须与滚动网格样式中的内边距保持一致，供缩放画布补偿固定标题位置。 */
const XLSX_SHEET_GRID_PADDING = 16;

/** 将 CSS 尺寸文本转换为可参与边框盒计算的像素值。 */
function readCssPixelValue(value: string) {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

/** 从滚动容器客户区中扣除内边距，得到不包含滚动槽的工作表可用视口。 */
function readGridViewportSize(grid: HTMLDivElement): XlsxSheetViewportSize {
  const style = window.getComputedStyle(grid);
  const horizontalInset =
    readCssPixelValue(style.paddingLeft) +
    readCssPixelValue(style.paddingRight);
  const verticalInset =
    readCssPixelValue(style.paddingTop) +
    readCssPixelValue(style.paddingBottom);
  return {
    width: Math.max(
      0,
      Math.round((grid.clientWidth - horizontalInset) * 100) / 100,
    ),
    height: Math.max(
      0,
      Math.round((grid.clientHeight - verticalInset) * 100) / 100,
    ),
  };
}

/** 渲染支持大数据窗口化的工作表网格。 */
function XlsxSheetGridComponent({ sheet, zoom, viewMode }: XlsxSheetGridProps) {
  const scale = zoom / 100;
  // 内边距属于未缩放的滚动容器，固定层位于缩放画布内，因此需要换算回逻辑像素。
  const stickyInset = -XLSX_SHEET_GRID_PADDING / Math.max(scale, 0.01);
  const gridRef = useRef<HTMLDivElement>(null);
  const tableRef = useRef<HTMLTableElement>(null);
  const [viewportSize, setViewportSize] =
    useState<XlsxSheetViewportSize>(EMPTY_VIEWPORT_SIZE);

  // 补位表格脱离文档流后，边框盒只会随外部布局变化，不再被补位行列反向撑大。
  useEffect(() => {
    const grid = gridRef.current;
    if (!grid) return undefined;
    const updateViewportSize = () => {
      const nextSize = readGridViewportSize(grid);
      setViewportSize((currentSize) =>
        currentSize.width === nextSize.width &&
        currentSize.height === nextSize.height
          ? currentSize
          : nextSize,
      );
    };
    updateViewportSize();

    if (typeof ResizeObserver === 'undefined') {
      window.addEventListener('resize', updateViewportSize);
      return () => window.removeEventListener('resize', updateViewportSize);
    }

    const observer = new ResizeObserver(updateViewportSize);
    try {
      observer.observe(grid, { box: 'border-box' });
    } catch {
      // 较旧实现可能支持 ResizeObserver 但不接受 box 参数，此时退回默认内容盒监听。
      observer.observe(grid);
    }
    return () => observer.disconnect();
  }, []);

  const sourceMetrics = useMemo(
    () =>
      getXlsxSheetMetrics(
        sheet,
        viewportSize.width,
        viewportSize.height,
        scale,
      ),
    [scale, sheet, viewportSize.height, viewportSize.width],
  );
  const measuredLayout = useXlsxSheetTableLayout(tableRef, sheet);
  const metrics = useMemo(
    () =>
      getXlsxSheetMetrics(
        sheet,
        viewportSize.width,
        viewportSize.height,
        scale,
        measuredLayout,
      ),
    [measuredLayout, scale, sheet, viewportSize.height, viewportSize.width],
  );
  // 这里使用 zoom 是为了让表格、图片和图表保持同一个坐标系缩放。
  const canvasStyle = useMemo<CSSProperties>(
    () => ({
      width: metrics.canvasWidth,
      minWidth: metrics.canvasWidth,
      minHeight: metrics.canvasHeight,
      zoom: scale,
    }),
    [metrics.canvasHeight, metrics.canvasWidth, scale],
  );

  return (
    <div ref={gridRef} className="office-file-xlsx-sheet-grid">
      <div className="office-file-xlsx-sheet-grid__canvas" style={canvasStyle}>
        <XlsxSheetFiller
          sheet={sheet}
          metrics={metrics}
          stickyInset={stickyInset}
        />
        <XlsxSheetTable
          sheet={sheet}
          tableWidth={sourceMetrics.tableWidth}
          visibleColumnWidths={sourceMetrics.visibleColumnWidths}
          visibleRowHeights={sourceMetrics.visibleRowHeights}
          renderedColumnHeaderHeight={metrics.columnHeaderHeight}
          renderedRowHeaderHeights={metrics.visibleRowHeights}
          stickyInset={stickyInset}
          tableRef={tableRef}
          viewMode={viewMode}
        />
        <XlsxFloatingImages sheet={sheet} metrics={metrics} />
        <XlsxFloatingCharts sheet={sheet} metrics={metrics} />
      </div>
    </div>
  );
}

export const XlsxSheetGrid = memo(XlsxSheetGridComponent);
