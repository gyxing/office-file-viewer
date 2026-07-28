// sheetRenderUtils 提供 XLSX 工作表渲染所需的样式转换和画布尺寸计算。
import type { CSSProperties } from 'react';
import type {
  XlsxAnchorPoint,
  XlsxCell,
  XlsxCellStyle,
  XlsxSheet,
} from '../../services/xlsx/types';

/** 工作表左侧行号栏的固定宽度。 */
export const XLSX_ROW_HEADER_WIDTH = 48;
/** 工作表顶部列标题栏的固定高度。 */
export const XLSX_COLUMN_HEADER_HEIGHT = 28;
/** 旧工作表模型未携带默认列宽时使用的兼容值。 */
export const XLSX_FALLBACK_COLUMN_WIDTH = 64;
/** 旧工作表模型未携带默认行高时使用的兼容值。 */
export const XLSX_FALLBACK_ROW_HEIGHT = 20;
/** 最后一行存在视觉内容时，底部至少保留的空白网格行数。 */
const XLSX_MIN_TRAILING_FILLER_ROWS = 2;

/** 描述浏览器完成表格布局后得到的真实尺寸。 */
export type XlsxMeasuredTableLayout = {
  /** 包含自动撑高内容和边框的完整表格高度。 */
  tableHeight: number;
  /** 浏览器最终计算出的列标题行高度。 */
  columnHeaderHeight: number;
  /** 过滤隐藏行后，各数据行由浏览器最终计算出的高度。 */
  visibleRowHeights: number[];
};

/** 描述工作表在当前视口中的真实尺寸和显示补位数量。 */
export type XlsxSheetMetrics = {
  /** 源工作表中可见列的标准化宽度，顺序与过滤隐藏列后的列模型一致。 */
  visibleColumnWidths: number[];
  /** 源工作表中可见行的标准化高度，顺序与过滤隐藏行后的行模型一致。 */
  visibleRowHeights: number[];
  /** 列标题行在当前表格布局中的真实高度。 */
  columnHeaderHeight: number;
  /** 源工作表可见行列构成的表格宽度。 */
  tableWidth: number;
  /** 源工作表可见行列构成的表格高度。 */
  tableHeight: number;
  /** 同时覆盖真实表格与当前视口的补位显示区宽度。 */
  renderedTableWidth: number;
  /** 同时覆盖真实表格与当前视口的补位显示区高度。 */
  renderedTableHeight: number;
  /** 同时覆盖表格、图片和图表的画布宽度。 */
  canvasWidth: number;
  /** 同时覆盖表格、图片和图表的画布高度。 */
  canvasHeight: number;
  /** 为覆盖当前视口而补充的显示列数量。 */
  fillerColumnCount: number;
  /** 为覆盖当前视口或保留末尾间距而补充的显示行数量。 */
  fillerRowCount: number;
  /** 补充列使用的工作表默认宽度。 */
  defaultColumnWidth: number;
  /** 补充行使用的工作表默认高度。 */
  defaultRowHeight: number;
};

// 单元格字体、填充、边框等来自工作簿样式表，只能在运行时转成 CSS，不能放到 Less。
/** 根据输入构建 `buildXlsxCellStyle` 返回的标准化结果。 */
export function buildXlsxCellStyle(cell: XlsxCell): CSSProperties {
  const style = cell.style ?? {};
  const css: CSSProperties = {
    fontWeight: style.bold ? 700 : 400,
    fontStyle: style.italic ? 'italic' : undefined,
    textDecoration: style.underline ? 'underline' : undefined,
    color: style.color,
    background: style.backgroundColor,
    textAlign: style.horizontalAlign,
    verticalAlign: style.verticalAlign,
    fontFamily: style.fontFamily,
    fontSize: style.fontSize,
    whiteSpace: style.wrapText ? 'pre-wrap' : 'nowrap',
    overflowWrap: style.wrapText ? 'anywhere' : undefined,
    wordBreak: style.wrapText ? 'break-word' : undefined,
  };
  return Object.fromEntries(
    Object.entries(css).filter(([, value]) => value !== undefined),
  ) as CSSProperties;
}

/** 判断 `isHighlightedXlsxCell` 对应的条件是否成立。 */
export function isHighlightedXlsxCell(style?: XlsxCellStyle) {
  return Boolean(style?.color?.toLowerCase() === '#ff0000' || style?.bold);
}

/** 将从 1 开始的列索引转换为 Excel 列标。 */
export function getSpreadsheetColumnLabel(columnIndex: number) {
  let remaining = Number.isFinite(columnIndex)
    ? Math.max(1, Math.floor(columnIndex))
    : 1;
  let label = '';
  while (remaining > 0) {
    const offset = (remaining - 1) % 26;
    label = String.fromCharCode(65 + offset) + label;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return label;
}

/** 保证来源文件中的尺寸异常时仍能得到可渲染的正数。 */
function normalizeSheetDimension(value: number | undefined, fallback: number) {
  return Number.isFinite(value) && Number(value) > 0 ? Number(value) : fallback;
}

/** 描述可按浏览器实测尺寸重新定位的工作表浮动对象锚点。 */
type XlsxFloatingAnchor = {
  /** 浮动对象的起始锚点。 */
  from: XlsxAnchorPoint;
  /** 浮动对象的结束锚点。 */
  to: XlsxAnchorPoint;
};

/** 描述浮动对象在工作表数据区中的实测矩形。 */
export type XlsxMeasuredAnchorRect = {
  /** 相对工作表数据区左侧的横坐标。 */
  x: number;
  /** 相对工作表数据区顶部的纵坐标。 */
  y: number;
  /** 按实测行列尺寸换算后的宽度。 */
  width: number;
  /** 按实测行列尺寸换算后的高度。 */
  height: number;
};

/** 按可见行列的实测尺寸换算单个锚点轴坐标。 */
function resolveMeasuredAnchorAxis(
  items: Array<{ index: number; size: number; hidden?: boolean }>,
  measuredSizes: number[],
  targetIndex: number,
  sourceOffset: number,
  fallbackSize: number,
) {
  let position = 0;
  let visibleIndex = 0;
  for (const item of items) {
    if (item.hidden) continue;
    const sourceSize = normalizeSheetDimension(item.size, fallbackSize);
    const measuredSize = normalizeSheetDimension(
      measuredSizes[visibleIndex],
      sourceSize,
    );
    if (item.index === targetIndex) {
      return position + measuredSize * (sourceOffset / sourceSize);
    }
    if (item.index > targetIndex) return position;
    position += measuredSize;
    visibleIndex += 1;
  }
  return position;
}

/** 按浏览器最终表格行高重算浮动图片或图表的锚点矩形。 */
export function getXlsxMeasuredAnchorRect(
  sheet: XlsxSheet,
  metrics: XlsxSheetMetrics,
  anchor: XlsxFloatingAnchor,
): XlsxMeasuredAnchorRect {
  const columns = sheet.columns.map((column) => ({
    index: column.index,
    size: column.width,
    hidden: column.hidden,
  }));
  const rows = sheet.rows.map((row) => ({
    index: row.index,
    size: row.height,
    hidden: row.hidden,
  }));
  const fromX = resolveMeasuredAnchorAxis(
    columns,
    metrics.visibleColumnWidths,
    anchor.from.column,
    anchor.from.columnOffset,
    metrics.defaultColumnWidth,
  );
  const toX = resolveMeasuredAnchorAxis(
    columns,
    metrics.visibleColumnWidths,
    anchor.to.column,
    anchor.to.columnOffset,
    metrics.defaultColumnWidth,
  );
  const fromY = resolveMeasuredAnchorAxis(
    rows,
    metrics.visibleRowHeights,
    anchor.from.row,
    anchor.from.rowOffset,
    metrics.defaultRowHeight,
  );
  const toY = resolveMeasuredAnchorAxis(
    rows,
    metrics.visibleRowHeights,
    anchor.to.row,
    anchor.to.rowOffset,
    metrics.defaultRowHeight,
  );
  return {
    x: fromX,
    y: fromY,
    width: Math.max(1, toX - fromX),
    height: Math.max(1, toY - fromY),
  };
}

/** 判断工作表最后一个可见行是否包含会实际显示的内容或格式。 */
function hasVisualContentInLastVisibleRow(sheet: XlsxSheet) {
  const lastVisibleRow = [...sheet.rows].reverse().find((row) => !row.hidden);
  if (!lastVisibleRow) return false;

  const hasVisibleCell = lastVisibleRow.cells.some((cell) => {
    if (cell.hiddenByMerge) return false;
    const style = cell.style;
    const hasVisibleStyle = Boolean(
      style?.backgroundColor ||
        style?.border ||
        style?.borderTop ||
        style?.borderRight ||
        style?.borderBottom ||
        style?.borderLeft,
    );
    return Boolean(
      cell.value || cell.formula || cell.formulaTokens || hasVisibleStyle,
    );
  });
  if (hasVisibleCell) return true;

  // 合并区域的非锚点单元格不会直接渲染，仍需按区域是否覆盖末行判断视觉占位。
  return sheet.merges.some(
    (merge) =>
      merge.startRow <= lastVisibleRow.index &&
      merge.endRow >= lastVisibleRow.index,
  );
}

// 画布需要同时覆盖单元格区域和浮动图片/图表，否则滚动区域会截断绝对定位元素。
/** 根据源工作表、视口和缩放比例计算真实尺寸及显示补位数量。 */
export function getXlsxSheetMetrics(
  sheet: XlsxSheet,
  viewportWidth = 0,
  viewportHeight = 0,
  scale = 1,
  measuredLayout?: XlsxMeasuredTableLayout,
): XlsxSheetMetrics {
  const defaultColumnWidth = normalizeSheetDimension(
    sheet.defaultColumnWidth,
    XLSX_FALLBACK_COLUMN_WIDTH,
  );
  const defaultRowHeight = normalizeSheetDimension(
    sheet.defaultRowHeight,
    XLSX_FALLBACK_ROW_HEIGHT,
  );
  const safeScale = Number.isFinite(scale) && scale > 0 ? scale : 1;
  const visibleColumnWidths = sheet.columns.flatMap((column) =>
    column.hidden
      ? []
      : [normalizeSheetDimension(column.width, defaultColumnWidth)],
  );
  const sourceVisibleRowHeights = sheet.rows.flatMap((row) =>
    row.hidden ? [] : [normalizeSheetDimension(row.height, defaultRowHeight)],
  );
  // 只有行数完全一致时才采纳实测布局，避免渐进解析或切换工作表时误用旧尺寸。
  const hasCompatibleMeasurement =
    measuredLayout?.visibleRowHeights.length === sourceVisibleRowHeights.length;
  const visibleRowHeights = hasCompatibleMeasurement
    ? sourceVisibleRowHeights.map((fallback, index) =>
        normalizeSheetDimension(
          measuredLayout?.visibleRowHeights[index],
          fallback,
        ),
      )
    : sourceVisibleRowHeights;
  const columnHeaderHeight = hasCompatibleMeasurement
    ? normalizeSheetDimension(
        measuredLayout?.columnHeaderHeight,
        XLSX_COLUMN_HEADER_HEIGHT,
      )
    : XLSX_COLUMN_HEADER_HEIGHT;
  const tableWidth =
    XLSX_ROW_HEADER_WIDTH +
    visibleColumnWidths.reduce((sum, width) => sum + width, 0);
  const calculatedTableHeight =
    columnHeaderHeight +
    visibleRowHeights.reduce((sum, height) => sum + height, 0);
  const tableHeight = hasCompatibleMeasurement
    ? Math.max(
        calculatedTableHeight,
        normalizeSheetDimension(
          measuredLayout?.tableHeight,
          calculatedTableHeight,
        ),
      )
    : calculatedTableHeight;
  const logicalViewportWidth =
    Number.isFinite(viewportWidth) && viewportWidth > 0
      ? viewportWidth / safeScale
      : 0;
  const logicalViewportHeight =
    Number.isFinite(viewportHeight) && viewportHeight > 0
      ? viewportHeight / safeScale
      : 0;
  const fillerColumnCount = Math.max(
    0,
    Math.ceil((logicalViewportWidth - tableWidth) / defaultColumnWidth),
  );
  const minimumTrailingRowCount = hasVisualContentInLastVisibleRow(sheet)
    ? XLSX_MIN_TRAILING_FILLER_ROWS
    : 0;
  const fillerRowCount = Math.max(
    minimumTrailingRowCount,
    Math.ceil((logicalViewportHeight - tableHeight) / defaultRowHeight),
  );
  // 补位数量向上取整以生成末尾半格的标题和网格，但显示范围必须精确止于视口边缘，
  // 否则最后一个完整补位格会反向撑大滚动区域，制造并非由真实内容引起的滚动条。
  const renderedTableWidth = Math.max(tableWidth, logicalViewportWidth);
  const renderedTableHeight = Math.max(
    tableHeight + minimumTrailingRowCount * defaultRowHeight,
    logicalViewportHeight,
  );
  return {
    visibleColumnWidths,
    visibleRowHeights,
    columnHeaderHeight,
    tableWidth,
    tableHeight,
    renderedTableWidth,
    renderedTableHeight,
    canvasWidth: Math.max(
      tableWidth,
      ...sheet.images.map(
        (image) => XLSX_ROW_HEADER_WIDTH + image.x + image.width,
      ),
      ...sheet.charts.map(
        (chart) => XLSX_ROW_HEADER_WIDTH + chart.x + chart.width,
      ),
    ),
    canvasHeight: Math.max(
      tableHeight,
      ...sheet.images.map(
        (image) => XLSX_COLUMN_HEADER_HEIGHT + image.y + image.height,
      ),
      ...sheet.charts.map(
        (chart) => XLSX_COLUMN_HEADER_HEIGHT + chart.y + chart.height,
      ),
    ),
    fillerColumnCount,
    fillerRowCount,
    defaultColumnWidth,
    defaultRowHeight,
  };
}
