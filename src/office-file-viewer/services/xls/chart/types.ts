import type {
  OfficeChartModel,
  OfficeChartType,
  OfficeDataLabels,
} from '../../../shared/ooxml/charts';
import type { SpreadsheetWarning } from '../../spreadsheet/types';
import type { Biff8Record } from '../biff8/Biff8Reader';
import type { Biff8Anchor } from '../drawing/types';
import type {
  Biff8ChartSubstream,
  Biff8SheetDescriptor,
  Biff8Workbook,
  Biff8Worksheet,
} from '../types';

/** 按 BEGIN/END 层级组织的 BIFF8 图表记录。 */
export type Biff8ChartRecordNode = {
  /** 在所属集合中的唯一标识。 */
  id: number;
  /** 在所属数据范围中的偏移位置。 */
  offset: number;
  /** 当前图表记录正文的原始字节。 */
  data: Uint8Array;
  /** 位于当前 BEGIN/END 容器内的子记录。 */
  children: Biff8ChartRecordNode[];
};

/** BIFF8 图表中的单个数据系列。 */
export type Biff8ChartSeries = {
  /** 面向用户展示的名称。 */
  name: string;
  /** 数据系列所属图表组的零基索引。 */
  groupIndex: number;
  /** 当前数据系列采用的标准图表类型。 */
  type?: OfficeChartType;
  /** 与数据点一一对应的分类轴标签。 */
  categories: Array<string | number | null>;
  /** 按数据点顺序排列的系列数值。 */
  values: Array<number | null>;
  /** 气泡图各数据点对应的气泡大小。 */
  bubbleSizes: Array<number | null>;
  /** 图表数据标签显示配置。 */
  dataLabels?: OfficeDataLabels;
  /** 数据系列的堆积方式。 */
  stacking?: 'stacked' | 'percentStacked';
  /** 前景或文字颜色，使用 CSS 颜色值。 */
  color?: string;
  /** 数据系列的数据点标记样式；未提供时不绘制标记。 */
  marker?: {
    /** 数据点标记使用的图形名称。 */
    symbol?: string;
    /** 当前数据占用的空间大小。 */
    size?: number;
  };
  /** 数据系列线条宽度，单位为标准化像素。 */
  lineWidth?: number;
};

/** 从 BIFF8 图表记录恢复的中间图表模型。 */
export type Biff8ChartModel = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 来源 BIFF8 图表的主类型名称。 */
  sourceType: string;
  /** 组合图表中各图表组的类型名称。 */
  groupTypes: string[];
  /** 图表是否启用三维显示。 */
  is3d: boolean;
  /** 图表是否包含次坐标轴。 */
  hasSecondaryAxis: boolean;
  /** 面向用户展示的标题。 */
  title?: string;
  /** 图表分类轴使用的标签。 */
  categories: string[];
  /** 按绘制顺序排列的数据系列。 */
  series: Biff8ChartSeries[];
  /** 是否显示图表图例。 */
  showLegend: boolean;
  /** 图例停靠位置。 */
  legendPosition?: 'top' | 'bottom' | 'left' | 'right';
  /** 相邻分类组之间的间距百分比。 */
  gapWidth?: number;
  /** 同一分类中各数据系列的重叠百分比。 */
  overlap?: number;
  /** 环形图中心孔径占图表直径的百分比。 */
  holeSize?: number;
  /** 对象在工作表或画布中的定位锚点。 */
  anchor: Biff8Anchor;
  /** 图表静态预览图片的资源地址。 */
  previewImageSrc?: string;
  /** 解析时产生但不阻止继续预览的警告。 */
  warnings: SpreadsheetWarning[];
};

/** 汇总XLS/BIFF8 解析当前步骤需要共享的上下文。 */
export type Biff8ChartContext = {
  /** 对应的 BIFF8 图表子流。 */
  substream: Biff8ChartSubstream;
  /** 当前处理的标准化工作簿。 */
  workbook: Biff8Workbook;
  /** 解析公式引用时使用的来源工作表。 */
  sourceSheet?: Biff8Worksheet;
  /** 当前项目的轻量描述信息。 */
  descriptor: Biff8SheetDescriptor;
  /** 图表在所属图表集合中的索引。 */
  chartIndex: number;
  /** 对象在工作表或画布中的定位锚点。 */
  anchor?: Biff8Anchor;
  /** 图表静态预览图片的资源地址。 */
  previewImageSrc?: string;
};

/** BIFF8 图表转换后的模型和渲染策略。 */
export type AdaptedBiff8Chart = {
  /** 转换后的标准图表模型。 */
  chart: OfficeChartModel;
  /** 图表最终使用交互渲染还是静态快照。 */
  renderMode: 'interactive' | 'snapshot';
  /** 触发静态降级的原始图表类型。 */
  degradedFrom?: string;
  /** 解析时产生但不阻止继续预览的警告。 */
  warnings: SpreadsheetWarning[];
};

/** 按工作表、行和列索引的图表数据缓存。 */
export type Biff8ChartCache = Map<
  number,
  Map<number, Map<number, string | number | boolean | null>>
>;

/** BIFF8 图表子流及其记录集合。 */
export type ChartStreamReadResult = {
  /** 对应的 BIFF8 图表子流。 */
  substream: Biff8ChartSubstream;
  /** 当前图表子流解析出的全部 BIFF8 记录。 */
  records: Biff8Record[];
};
