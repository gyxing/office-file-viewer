// useXlsxSheetTableLayout 读取浏览器实际表格布局，统一校准补位层和浮动对象锚点。
import type { RefObject } from 'react';
import { useEffect, useState } from 'react';
import type { XlsxSheet } from '../../services/xlsx/types';
import type { XlsxMeasuredTableLayout } from './sheetRenderUtils';

/** 记录某个工作表引用对应的最近一次表格测量结果。 */
type XlsxSheetTableLayoutState = {
  /** 产生当前测量结果的工作表对象。 */
  sheet: XlsxSheet;
  /** 浏览器最终计算出的表格布局。 */
  layout: XlsxMeasuredTableLayout;
};

/** 判断两组行高是否完全一致，避免 ResizeObserver 回调引起无效渲染。 */
function areRowHeightsEqual(current: number[], next: number[]) {
  return (
    current.length === next.length &&
    current.every((height, index) => height === next[index])
  );
}

/** 将浏览器亚像素测量值归一化，避免 ResizeObserver 因浮点抖动重复更新。 */
function normalizeMeasuredSize(value: number) {
  return Math.round(value * 1000) / 1000;
}

/** 从真实 table、表头和数据行节点读取未受 CSS zoom 影响的布局尺寸。 */
function readTableLayout(table: HTMLTableElement): XlsxMeasuredTableLayout {
  const tableRect = table.getBoundingClientRect();
  const renderedScale =
    table.offsetWidth > 0 ? tableRect.width / table.offsetWidth : 1;
  const logicalScale =
    Number.isFinite(renderedScale) && renderedScale > 0 ? renderedScale : 1;
  const readLogicalHeight = (element: HTMLElement | undefined) =>
    element
      ? normalizeMeasuredSize(
          element.getBoundingClientRect().height / logicalScale,
        )
      : 0;
  const headerRow = table.tHead?.rows.item(0) ?? undefined;
  const body = table.tBodies.item(0);
  const visibleRowHeights = body
    ? Array.from(body.rows, (row) => readLogicalHeight(row))
    : [];
  return {
    tableHeight: normalizeMeasuredSize(tableRect.height / logicalScale),
    columnHeaderHeight: readLogicalHeight(headerRow),
    visibleRowHeights,
  };
}

/**
 * 持续读取当前工作表真实表格的布局结果。
 *
 * 表格中的换行文本和合并单元格可能把行高撑到源文件记录值以上，因此补位层与
 * 浮动对象都必须使用浏览器最终布局，不能只依赖解析模型。
 */
export function useXlsxSheetTableLayout(
  tableRef: RefObject<HTMLTableElement>,
  sheet: XlsxSheet,
) {
  const [state, setState] = useState<XlsxSheetTableLayoutState>();

  useEffect(() => {
    const table = tableRef.current;
    if (!table) return undefined;

    const updateLayout = () => {
      const nextLayout = readTableLayout(table);
      setState((current) => {
        if (
          current?.sheet === sheet &&
          current.layout.tableHeight === nextLayout.tableHeight &&
          current.layout.columnHeaderHeight === nextLayout.columnHeaderHeight &&
          areRowHeightsEqual(
            current.layout.visibleRowHeights,
            nextLayout.visibleRowHeights,
          )
        ) {
          return current;
        }
        return { sheet, layout: nextLayout };
      });
    };
    updateLayout();

    if (typeof ResizeObserver === 'undefined') {
      window.addEventListener('resize', updateLayout);
      return () => window.removeEventListener('resize', updateLayout);
    }

    const observer = new ResizeObserver(updateLayout);
    observer.observe(table);
    return () => observer.disconnect();
  }, [sheet, tableRef]);

  return state?.sheet === sheet ? state.layout : undefined;
}
