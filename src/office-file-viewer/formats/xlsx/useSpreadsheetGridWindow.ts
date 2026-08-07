import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { createSpreadsheetAxisIndex } from '../../services/spreadsheet/SpreadsheetAxisIndex';
import type { SpreadsheetPerformanceProfile } from '../../services/spreadsheet/spreadsheetPerformance';
import type { SpreadsheetSheetLayout } from '../../services/spreadsheet/SpreadsheetSource';
import type { SpreadsheetRange } from '../../services/spreadsheet/types';
import type { SpreadsheetViewMode } from '../../services/spreadsheet/viewMode';
import {
  XLSX_COLUMN_HEADER_HEIGHT,
  XLSX_ROW_HEADER_WIDTH,
} from './sheetRenderUtils';
import { mergeSpreadsheetReadingRowMetrics } from './spreadsheetReadingLayout';

/** 电子表格滚动视口的位置和尺寸。 */
type ViewportState = {
  /** 视口水平方向的滚动距离。 */
  scrollLeft: number;
  /** 视口垂直方向的滚动距离。 */
  scrollTop: number;
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
};

/** 电子表格网格尚未测量时使用的空视口。 */
const EMPTY_VIEWPORT: ViewportState = {
  scrollLeft: 0,
  scrollTop: 0,
  width: 0,
  height: 0,
};

/** SSR 环境延后布局副作用，浏览器中则在绘制前校正滚动锚点。 */
const useIsomorphicLayoutEffect =
  typeof window === 'undefined' ? useEffect : useLayoutEffect;

/** 虚拟工作表窗口计算所需的输入。 */
type UseSpreadsheetGridWindowOptions = {
  /** 当前工作表的稳定标识。 */
  sheetId: string;
  /** 当前工作表的轻量布局。 */
  layout: SpreadsheetSheetLayout;
  /** 当前工作表使用的窗口化策略。 */
  gridMode: SpreadsheetPerformanceProfile['gridMode'];
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 当前电子表格采用的显示模式。 */
  viewMode: SpreadsheetViewMode;
  /** 当前工作表已经计算出的稀疏阅读行高。 */
  readingRowHeights: ReadonlyMap<number, number>;
};

/** 根据二维滚动视口计算带约两个视口缓冲的全局行列范围。 */
export function useSpreadsheetGridWindow({
  sheetId,
  layout,
  gridMode,
  zoom,
  viewMode,
  readingRowHeights,
}: UseSpreadsheetGridWindowOptions) {
  const viewportRef = useRef<HTMLDivElement>(null);
  const [viewport, setViewport] = useState(EMPTY_VIEWPORT);
  const scale = Math.max(0.01, zoom / 100);
  const sourceRowAxis = useMemo(
    () =>
      createSpreadsheetAxisIndex(
        layout.rowCount,
        layout.defaultRowHeight,
        layout.rows,
      ),
    [layout],
  );
  const rowAxis = useMemo(() => {
    if (viewMode === 'source') return sourceRowAxis;
    return createSpreadsheetAxisIndex(
      layout.rowCount,
      layout.defaultRowHeight,
      mergeSpreadsheetReadingRowMetrics(layout.rows, readingRowHeights),
    );
  }, [layout, readingRowHeights, sourceRowAxis, viewMode]);
  const columnAxis = useMemo(
    () =>
      createSpreadsheetAxisIndex(
        layout.columnCount,
        layout.defaultColumnWidth,
        layout.columns,
      ),
    [layout],
  );

  useEffect(() => {
    const element = viewportRef.current;
    if (!element) return undefined;
    const update = () => {
      const next = {
        scrollLeft: element.scrollLeft,
        scrollTop: element.scrollTop,
        width: element.clientWidth,
        height: element.clientHeight,
      };
      setViewport((current) =>
        current.scrollLeft === next.scrollLeft &&
        current.scrollTop === next.scrollTop &&
        current.width === next.width &&
        current.height === next.height
          ? current
          : next,
      );
    };
    update();
    element.addEventListener('scroll', update, { passive: true });
    const observer =
      typeof ResizeObserver === 'undefined'
        ? undefined
        : new ResizeObserver(update);
    observer?.observe(element);
    window.addEventListener('resize', update);
    return () => {
      element.removeEventListener('scroll', update);
      observer?.disconnect();
      window.removeEventListener('resize', update);
    };
  }, []);

  const previousRowAxisRef = useRef({ sheetId, rowAxis });
  useIsomorphicLayoutEffect(() => {
    const previous = previousRowAxisRef.current;
    previousRowAxisRef.current = { sheetId, rowAxis };
    const element = viewportRef.current;
    if (
      !element ||
      previous.sheetId !== sheetId ||
      previous.rowAxis === rowAxis ||
      element.scrollTop <= 0
    ) {
      return;
    }

    // 固定列标题在画布内随 scrollTop 反向定位，数据区可视起点实际对应 scrollTop 本身。
    const previousLogicalTop = element.scrollTop / scale;
    const anchorRow = previous.rowAxis.findIndexAtOffset(previousLogicalTop);
    const offsetWithinRow =
      previousLogicalTop - previous.rowAxis.offsetAt(anchorRow);
    const nextLogicalTop =
      rowAxis.offsetAt(anchorRow) +
      Math.min(Math.max(0, offsetWithinRow), rowAxis.sizeAt(anchorRow));
    const nextScrollTop = nextLogicalTop * scale;
    if (Math.abs(element.scrollTop - nextScrollTop) > 0.5) {
      element.scrollTop = nextScrollTop;
    }
  }, [rowAxis, scale, sheetId]);

  const range = useMemo<SpreadsheetRange>(() => {
    const logicalLeft = Math.max(
      0,
      viewport.scrollLeft / scale - XLSX_ROW_HEADER_WIDTH,
    );
    const logicalTop = Math.max(
      0,
      viewport.scrollTop / scale - XLSX_COLUMN_HEADER_HEIGHT,
    );
    const logicalWidth = Math.max(320, viewport.width / scale);
    const logicalHeight = Math.max(240, viewport.height / scale);
    const startRow = rowAxis.findIndexAtOffset(
      Math.max(0, logicalTop - logicalHeight * 2),
    );
    const endRow = rowAxis.findIndexAtOffset(logicalTop + logicalHeight * 3);
    const startColumn =
      gridMode === 'row-window'
        ? 1
        : columnAxis.findIndexAtOffset(
            Math.max(0, logicalLeft - logicalWidth * 2),
          );
    const endColumn =
      gridMode === 'row-window'
        ? layout.columnCount
        : columnAxis.findIndexAtOffset(logicalLeft + logicalWidth * 3);
    return {
      startRow: Math.max(1, startRow),
      endRow: Math.min(layout.rowCount, Math.max(startRow, endRow)),
      startColumn: Math.max(1, startColumn),
      endColumn: Math.min(layout.columnCount, Math.max(startColumn, endColumn)),
    };
  }, [
    columnAxis,
    gridMode,
    layout.columnCount,
    layout.rowCount,
    rowAxis,
    scale,
    viewport,
  ]);

  return {
    viewportRef,
    viewport,
    range,
    sourceRowAxis,
    rowAxis,
    columnAxis,
    scale,
    canvasWidth:
      XLSX_ROW_HEADER_WIDTH + columnAxis.offsetAt(layout.columnCount + 1),
    canvasHeight:
      XLSX_COLUMN_HEADER_HEIGHT + rowAxis.offsetAt(layout.rowCount + 1),
  };
}
