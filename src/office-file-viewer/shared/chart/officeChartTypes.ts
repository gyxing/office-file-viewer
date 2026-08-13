/** 图表渐变中的单个颜色停止点。 */
export type OfficeChartColorStop = {
  /** 在所属数据范围中的偏移位置。 */
  offset: number;
  /** 前景或文字颜色，使用 CSS 颜色值。 */
  color: string;
};

/** 图表支持的纯色或线性渐变颜色。 */
export type OfficeChartColor =
  | string
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'linear';
      /** 相对定位区域左侧的横坐标，单位为标准化渲染像素。 */
      x: number;
      /** 相对定位区域顶部的纵坐标，单位为标准化渲染像素。 */
      y: number;
      /** 渐变终点的横坐标。 */
      x2: number;
      /** 渐变终点的纵坐标。 */
      y2: number;
      /** 按偏移位置排列的渐变颜色停止点。 */
      colorStops: OfficeChartColorStop[];
      /** 渐变坐标是否相对于全局画布计算。 */
      global?: boolean;
    };

/** 标准图表模型支持的图表类型。 */
export type OfficeChartType =
  | 'line'
  | 'bar'
  | 'column'
  | 'pie'
  | 'doughnut'
  | 'area'
  | 'scatter'
  | 'bubble'
  | 'radar'
  | 'map'
  | 'unknown';

/** 标准图表模型中的单个数据系列。 */
export type OfficeChartSeries = {
  /** 面向用户展示的名称。 */
  name: string;
  /** 按数据点顺序排列的系列纵轴值。 */
  values: number[];
  /** 散点图或气泡图使用的横轴数值。 */
  xValues?: number[];
  /** 气泡图各数据点对应的气泡大小。 */
  bubbleSizes?: number[];
  /** 当前数据系列采用的标准图表类型。 */
  type?: OfficeChartType;
  /** 数据系列的堆积方式。 */
  stacking?: 'stacked' | 'percentStacked';
  /** 组合图表中参与同组堆积的标识。 */
  stackGroup?: string;
  /** 相邻分类组之间的间距百分比。 */
  gapWidth?: number;
  /** 同一分类中各数据系列的重叠百分比。 */
  overlap?: number;
  /** 前景或文字颜色，使用 CSS 颜色值。 */
  color?: string;
  /** 按数据点顺序覆盖的颜色。 */
  pointColors?: string[];
  /** 按数据点顺序覆盖的标签文本。 */
  pointLabels?: string[];
  /** 按数据点顺序覆盖的填充和边框样式。 */
  pointStyles?: Array<{
    /** 前景或文字颜色，使用 CSS 颜色值。 */
    color?: OfficeChartColor;
    /** 边框颜色。 */
    borderColor?: string;
    /** 边框宽度，单位为标准化渲染像素。 */
    borderWidth?: number;
  }>;
  /** 图表数据标签显示配置。 */
  dataLabels?: OfficeDataLabels;
  /** 是否使用平滑曲线连接数据点。 */
  smooth?: boolean;
  /** 数据系列线条宽度，单位为标准化像素。 */
  lineWidth?: number;
  /** 数据系列的数据点标记样式；未提供时不绘制标记。 */
  marker?: {
    /** 数据点标记使用的图形名称。 */
    symbol?: string;
    /** 当前数据占用的空间大小。 */
    size?: number;
  };
};

/** 图表数据标签的内容和显示规则。 */
export type OfficeDataLabels = {
  /** 是否按源文件指示隐藏当前图表元素。 */
  delete?: boolean;
  /** 对象的定位信息及其参考坐标系。 */
  position?: string;
  /** 同时显示多个标签值时使用的分隔文本。 */
  separator?: string;
  /** 数据标签是否显示对应系列的图例标记。 */
  showLegendKey?: boolean;
  /** 数据标签是否显示数据点数值。 */
  showVal?: boolean;
  /** 数据标签是否显示分类名称。 */
  showCatName?: boolean;
  /** 数据标签是否显示系列名称。 */
  showSerName?: boolean;
  /** 数据标签是否显示百分比。 */
  showPercent?: boolean;
  /** 数据标签是否显示气泡大小。 */
  showBubbleSize?: boolean;
  /** 数据标签是否显示引导线。 */
  showLeaderLines?: boolean;
};

/** 跨 PPT、PPTX、XLS 和 XLSX 共用的标准图表模型。 */
export type OfficeChartModel = {
  /** 图表采用的标准类型。 */
  type: OfficeChartType;
  /** 面向用户展示的标题。 */
  title?: string;
  /** 分类轴使用的标签。 */
  categories: string[];
  /** 按绘制顺序排列的数据系列。 */
  series: OfficeChartSeries[];
  /** 图表数据标签显示配置。 */
  dataLabels?: OfficeDataLabels;
  /** 是否显示图表图例。 */
  showLegend?: boolean;
  /** 图例停靠位置。 */
  legendPosition?: 'top' | 'bottom' | 'left' | 'right';
  /** 图例尺寸和文字样式。 */
  legendStyle?: {
    /** 图例标记区域宽度，单位为标准化像素。 */
    itemWidth?: number;
    /** 图例标记区域高度，单位为标准化像素。 */
    itemHeight?: number;
    /** 图例文本的字体与颜色样式；未提供时沿用图表主题。 */
    textStyle?: {
      /** 前景或文字颜色，使用 CSS 颜色值。 */
      color?: string;
      /** 字体族名称。 */
      fontFamily?: string;
      /** 字号，单位为标准化渲染像素。 */
      fontSize?: number;
      /** 字体样式。 */
      fontStyle?: 'normal' | 'italic' | 'oblique';
      /** 字体粗细。 */
      fontWeight?: 'normal' | 'bold' | 'bolder' | 'lighter' | number;
    };
  };
  /** 是否显示图表数据标签。 */
  showDataLabels?: boolean;
  /** 环形图中心孔径占图表直径的百分比。 */
  holeSize?: number;
  /** 饼图第一扇区的起始角度，单位为度。 */
  startAngle?: number;
  /** 复合饼图第二绘图区采用条形还是饼形。 */
  ofPieType?: 'bar' | 'pie';
  /** 复合饼图第二绘图区包含的数据点数量。 */
  ofPieSecondPlotCount?: number;
  /** 复合饼图第二绘图区相对主饼图的尺寸百分比。 */
  secondPieSize?: number;
  /** 相邻分类组之间的间距百分比。 */
  gapWidth?: number;
  /** 同一分类中各数据系列的重叠百分比。 */
  overlap?: number;
  /** 是否显示分类轴的轴线、刻度和标签。 */
  showCategoryAxis?: boolean;
  /** 是否显示数值轴的轴线、刻度和标签；网格线单独保留。 */
  showValueAxis?: boolean;
  /** 数值轴的显式最小值。 */
  valueAxisMinimum?: number;
  /** 数值轴的显式最大值。 */
  valueAxisMaximum?: number;
  /** 数值轴相邻主刻度的间隔。 */
  valueAxisMajorUnit?: number;
  /** 玫瑰图按半径还是面积编码数值。 */
  roseType?: 'radius' | 'area';
  /** 饼图或环形图使用的内外半径。 */
  radius?: [string, string];
  /** 雷达图的来源样式名称。 */
  radarStyle?: string;
  /** 雷达图相对容器的显示半径。 */
  radarRadius?: string;
  /** 雷达图第一个维度的起始角度，单位为度。 */
  radarStartAngle?: number;
  /** 雷达图半径轴划分的分段数量。 */
  radarSplitNumber?: number;
  /** 雷达图各维度的名称和最大值。 */
  radarIndicators?: Array<{
    /** 面向用户展示的名称。 */
    name: string;
    /** 雷达图当前维度允许的最大值。 */
    max: number;
  }>;
  /** 地图图表显示的数据系列名称。 */
  mapSeriesName?: string;
  /** 地图图表对应的地区名称。 */
  mapRegion?: string;
  /** 图表运行时注册的地图名称。 */
  mapName?: string;
  /** 地图图表按需加载的 GeoJSON 地址。 */
  mapGeoJsonUrl?: string;
  /** 交互图表不可用时使用的静态快照地址。 */
  snapshotSrc?: string;
  /** 来源文档中的原始图表类型。 */
  sourceType?: string;
  /** 图表使用交互渲染还是静态快照。 */
  renderMode?: 'interactive' | 'snapshot';
  /** 触发静态降级的原始图表类型。 */
  degradedFrom?: string;
};
