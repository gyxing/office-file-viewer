import { BIFF8_RECORD } from '../biff8/constants';

/** 图表组记录与通用图表类型的对应关系。 */
export const CHART_TYPE_RECORDS = new Map<number, string>([
  [BIFF8_RECORD.BAR, 'bar'],
  [BIFF8_RECORD.LINE, 'line'],
  [BIFF8_RECORD.PIE, 'pie'],
  [BIFF8_RECORD.AREA, 'area'],
  [BIFF8_RECORD.SCATTER, 'scatter'],
  [BIFF8_RECORD.RADAR, 'radar'],
  [BIFF8_RECORD.RADARAREA, 'radarArea'],
  [BIFF8_RECORD.SURF, 'surface'],
  [BIFF8_RECORD.BOPPOP, 'ofPie'],
]);

/** BIFF8 图表缓存数据类别对应的编号。 */
export const CHART_CACHE_KIND = {
  VALUES: 1,
  CATEGORIES: 2,
  BUBBLES: 3,
} as const;

/** BIFF8 图表文本对象链接目标对应的编号。 */
export const CHART_TEXT_LINK = {
  TITLE: 1,
  Y_AXIS_TITLE: 2,
  X_AXIS_TITLE: 3,
  SERIES_OR_POINT: 4,
} as const;

/** Excel 图表缺少锚点时使用的默认单元格范围。 */
export const FALLBACK_CHART_ANCHOR = {
  from: { row: 0, column: 0, rowFraction: 0, columnFraction: 0 },
  to: { row: 19, column: 9, rowFraction: 0, columnFraction: 0 },
} as const;
