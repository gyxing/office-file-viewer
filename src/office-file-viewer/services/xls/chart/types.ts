import type {
  OfficeChartModel,
  OfficeChartType,
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

/** 描述 Biff8ChartRecordNode 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8ChartRecordNode = {
  /** Biff8ChartRecordNode 在所属文档或任务中的唯一标识。 */
  id: number;
  /** Biff8ChartRecordNode 在源二进制流中的字节偏移。 */
  offset: number;
  /** Biff8ChartRecordNode 当前步骤需要处理的原始或标准化数据。 */
  data: Uint8Array;
  /** Biff8ChartRecordNode 包含并负责布局的 React 子节点。 */
  children: Biff8ChartRecordNode[];
};

/** 描述 Biff8ChartSeries 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8ChartSeries = {
  /** Biff8ChartSeries 的可读名称。 */
  name: string;
  /** Biff8ChartSeries 在所属集合中的位置索引。 */
  groupIndex: number;
  /** 用于区分 Biff8ChartSeries 不同结构分支的类型标识。 */
  type?: OfficeChartType;
  /** Biff8ChartSeries 包含的 categories 有序集合。 */
  categories: Array<string | number | null>;
  /** Biff8ChartSeries 包含的 values 有序集合。 */
  values: Array<number | null>;
  /** Biff8ChartSeries 包含的 bubbleSizes 有序集合。 */
  bubbleSizes: Array<number | null>;
  /** Biff8ChartSeries 关联的 stacking 结构；字段形状由 'stacked' | 'percentStacked' 定义；未提供时使用来源格式或渲染器的默认行为。 */
  stacking?: 'stacked' | 'percentStacked';
  /** Biff8ChartSeries 的前景或文本颜色，使用标准化 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
  color?: string;
  /** 数据系列的数据点标记样式；未提供时不绘制标记。 */
  marker?: {
    /** Biff8ChartSeries 的 symbol 文本值。 */
    symbol?: string;
    /** Biff8ChartSeries 对应二进制记录或数据块的字节长度。 */
    size?: number;
  };
  /** Biff8ChartSeries 的 lineWidth 渲染尺寸，单位为标准化像素；未提供时使用来源格式或渲染器的默认行为。 */
  lineWidth?: number;
};

/** 描述 XLS/BIFF8 解析使用的标准化模型。 */
export type Biff8ChartModel = {
  /** Biff8ChartModel 在所属文档或任务中的唯一标识。 */
  id: string;
  /** Biff8ChartModel 的 sourceType 文本值。 */
  sourceType: string;
  /** Biff8ChartModel 包含的 groupTypes 有序集合。 */
  groupTypes: string[];
  /** 图表是否启用三维显示。 */
  is3d: boolean;
  /** 图表是否包含次坐标轴。 */
  hasSecondaryAxis: boolean;
  /** Biff8ChartModel 对外展示的标题。 */
  title?: string;
  /** Biff8ChartModel 包含的 categories 有序集合。 */
  categories: string[];
  /** Biff8ChartModel 包含的 series 有序集合。 */
  series: Biff8ChartSeries[];
  /** 是否显示图表图例。 */
  showLegend: boolean;
  /** Biff8ChartModel 的图例停靠位置；未提供时使用来源格式或渲染器的默认行为。 */
  legendPosition?: 'top' | 'bottom' | 'left' | 'right';
  /** Biff8ChartModel 的 gapWidth 图表布局参数，数值语义遵循 Office 图表规范；未提供时使用来源格式或渲染器的默认行为。 */
  gapWidth?: number;
  /** Biff8ChartModel 的 overlap 图表布局参数，数值语义遵循 Office 图表规范；未提供时使用来源格式或渲染器的默认行为。 */
  overlap?: number;
  /** Biff8ChartModel 的 holeSize 图表布局参数，数值语义遵循 Office 图表规范；未提供时使用来源格式或渲染器的默认行为。 */
  holeSize?: number;
  /** Biff8ChartModel 在工作表或画布中的定位锚点。 */
  anchor: Biff8Anchor;
  /** Biff8ChartModel 的 previewImageSrc 文本值。 */
  previewImageSrc?: string;
  /** Biff8ChartModel 解析时产生但不阻止继续预览的警告集合。 */
  warnings: SpreadsheetWarning[];
};

/** 汇总XLS/BIFF8 解析当前步骤需要共享的上下文。 */
export type Biff8ChartContext = {
  /** Biff8ChartContext 对应的 BIFF8 图表子流。 */
  substream: Biff8ChartSubstream;
  /** Biff8ChartContext 当前关联的标准化工作簿。 */
  workbook: Biff8Workbook;
  /** Biff8ChartContext 关联的 sourceSheet 结构；字段形状由 Biff8Worksheet 定义；未提供时使用来源格式或渲染器的默认行为。 */
  sourceSheet?: Biff8Worksheet;
  /** Biff8ChartContext 关联的 descriptor 结构；字段形状由 Biff8SheetDescriptor 定义。 */
  descriptor: Biff8SheetDescriptor;
  /** 图表在所属图表集合中的索引。 */
  chartIndex: number;
  /** Biff8ChartContext 在工作表或画布中的定位锚点；未提供时使用来源格式或渲染器的默认行为。 */
  anchor?: Biff8Anchor;
  /** Biff8ChartContext 的 previewImageSrc 文本值。 */
  previewImageSrc?: string;
};

/** 描述 AdaptedBiff8Chart 在 XLS/BIFF8 解析中的数据结构。 */
export type AdaptedBiff8Chart = {
  /** AdaptedBiff8Chart 当前关联的图表模型。 */
  chart: OfficeChartModel;
  /** AdaptedBiff8Chart 关联的 renderMode 结构；字段形状由 'interactive' | 'snapshot' 定义。 */
  renderMode: 'interactive' | 'snapshot';
  /** AdaptedBiff8Chart 的 degradedFrom 文本值。 */
  degradedFrom?: string;
  /** AdaptedBiff8Chart 解析时产生但不阻止继续预览的警告集合。 */
  warnings: SpreadsheetWarning[];
};

/** 描述 Biff8ChartCache 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8ChartCache = Map<
  number,
  Map<number, Map<number, string | number | boolean | null>>
>;

/** 描述 XLS/BIFF8 解析产生的处理结果。 */
export type ChartStreamReadResult = {
  /** ChartStreamReadResult 对应的 BIFF8 图表子流。 */
  substream: Biff8ChartSubstream;
  /** ChartStreamReadResult 包含的 records 有序集合。 */
  records: Biff8Record[];
};
