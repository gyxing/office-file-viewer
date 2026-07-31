/** 描述演示文稿标准模型过程中可继续处理的警告。 */
export type PresentationWarning = {
  /** 供程序识别当前情况的稳定代码。 */
  code: string;
  /** 面向调用方或用户展示的说明。 */
  message: string;
  /** 幻灯片在演示文稿集合中的索引。 */
  slideIndex?: number;
  /** 产生警告的源对象标识。 */
  objectId?: string;
  /** 在所属数据范围中的偏移位置。 */
  offset?: number;
};

/** 记录演示文稿标准模型持有且需要统一管理的资源。 */
export type PresentationResources = {
  /** 浏览器创建的对象 URL 集合，文档释放时必须逐一撤销。 */
  objectUrls: string[];
};

/** 描述演示文稿标准模型生成的标准化文档模型。 */
export type PresentationDocument = {
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 当前文档使用的主题颜色和字体配置。 */
  theme: ThemeModel;
  /** 按演示文稿顺序排列的幻灯片。 */
  slides: SlideModel[];
  /** 解析时产生但不阻止继续预览的警告；未提供表示没有警告。 */
  warnings?: PresentationWarning[];
  /** 文档持有且需要在销毁时释放的浏览器资源；未提供表示没有需释放资源。 */
  resources?: PresentationResources;
};

export type {
  PresentationSlideDescriptor,
  PresentationSource,
  PresentationSourceSnapshot,
} from './PresentationSource';

/** 描述演示文稿标准模型使用的标准化模型。 */
export type ThemeModel = {
  /** 按主题槽名称索引的标准化颜色映射。 */
  colorScheme: Record<string, string>;
  /** 按主题字体槽名称索引的字体族映射。 */
  fontScheme: Record<string, string>;
  /** 将幻灯片颜色槽重定向到主题颜色槽的映射。 */
  colorMap?: Record<string, string>;
};

/** 渐变中的单个颜色停止点。 */
export type GradientStop = {
  /** 渐变停止点的位置比例，取值范围为 0 到 1。 */
  offset: number;
  /** 前景或文字颜色，使用 CSS 颜色值。 */
  color: string;
};

/** 线性渐变的角度和颜色停止点。 */
export type GradientFill = {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'linear';
  /** 角度值，单位为度。 */
  angle: number;
  /** 按位置排序的渐变颜色停止点。 */
  stops: GradientStop[];
};

/** 描述单页幻灯片关联的演讲者备注正文。 */
export type SpeakerNotesModel = {
  /** 按源文档顺序保留的备注段落。 */
  paragraphs: TextParagraph[];
  /** 供搜索、复制和无样式降级展示使用的纯文本。 */
  plainText: string;
};

/** 描述演示文稿标准模型使用的标准化模型。 */
export type SlideModel = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 幻灯片在演示文稿中的稳定顺序索引。 */
  index: number;
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 是否隐藏当前项目。 */
  hidden?: boolean;
  /** 当前页面、幻灯片或元素的背景配置。 */
  background?: SlideBackground;
  /** 当前幻灯片关联的演讲者备注正文；未提供表示没有可展示备注。 */
  speakerNotes?: SpeakerNotesModel;
  /** 按绘制顺序排列的演示文稿元素。 */
  elements: SlideElement[];
};

/** 幻灯片使用的颜色、渐变或图片背景。 */
export type SlideBackground = {
  /** 填充颜色、渐变或无填充标记。 */
  fill?: string;
  /** 填充区域的透明度，取值范围为 0 到 1。 */
  fillOpacity?: number;
  /** 用作幻灯片背景的媒体资源引用。 */
  imageRef?: string | OfficeResourceSource;
};

/** 描述演示文稿标准模型使用的样式参数。 */
export type TextStyle = {
  /** 字体族名称。 */
  fontFamily?: string;
  /** 字号，单位为标准化渲染像素。 */
  fontSize?: number;
  /** 是否使用粗体。 */
  bold?: boolean;
  /** 是否使用斜体。 */
  italic?: boolean;
  /** 是否绘制下划线。 */
  underline?: boolean;
  /** 使用的删除线类型。 */
  strike?: 'none' | 'sngStrike' | 'dblStrike';
  /** 是否将小写字母以小型大写字形显示。 */
  smallCaps?: boolean;
  /** 是否将文本全部以大写字形显示。 */
  allCaps?: boolean;
  /** 前景或文字颜色，使用 CSS 颜色值。 */
  color?: string;
  /** 文本字形使用的纯色或渐变填充；未提供时回退到 color。 */
  textFill?: string | GradientFill;
  /** 整体透明度，0 表示完全透明，1 表示完全不透明。 */
  opacity?: number;
  /** 水平对齐方式。 */
  align?: 'left' | 'center' | 'right' | 'justify';
  /** 垂直对齐方式。 */
  verticalAlign?: 'top' | 'middle' | 'bottom';
  /** 文本书写方向；未提供时采用水平从上到下。 */
  writingMode?: 'horizontal-tb' | 'vertical-rl' | 'vertical-lr';
  /** 文本溢出时采用不调整、缩小文本或扩大形状的策略。 */
  fit?: 'none' | 'shrinkText' | 'resizeShape';
  /** 左外边距，单位为标准化渲染像素。 */
  marginLeft?: number;
  /** 右外边距，单位为标准化渲染像素。 */
  marginRight?: number;
  /** 上外边距，单位为标准化渲染像素。 */
  marginTop?: number;
  /** 下外边距，单位为标准化渲染像素。 */
  marginBottom?: number;
  /** 行高，单位为标准化渲染像素。 */
  lineHeight?: number;
  /** 段前间距，单位为标准化渲染像素。 */
  spaceBefore?: number;
  /** 段后间距，单位为标准化渲染像素。 */
  spaceAfter?: number;
  /** 段落首行文本缩进，单位为标准化渲染像素。 */
  textIndent?: number;
  /** 字符间距，单位为标准化渲染像素。 */
  charSpace?: number;
  /** 文字相对基线的偏移，单位为标准化渲染像素。 */
  baseline?: number;
  /** 段落项目符号样式；未提供表示不覆盖段落级设置。 */
  bullet?: TextBulletStyle;
};

/** 描述演示文稿标准模型使用的样式参数。 */
export type TextBulletStyle = {
  /** 项目符号使用的字符。 */
  char?: string;
  /** 前景或文字颜色，使用 CSS 颜色值。 */
  color?: string;
  /** 项目符号字号，单位为标准化渲染像素。 */
  size?: number;
  /** 是否显式禁用项目符号。 */
  none?: boolean;
};

/** 具有统一文字样式的连续文本片段。 */
export type TextRun = {
  /** 文本内容。 */
  text: string;
  /** 当前内容使用的渲染样式。 */
  style?: TextStyle;
};

/** 由文本片段、段落样式和项目符号组成的段落。 */
export type TextParagraph = {
  /** 按显示顺序排列的连续文本片段。 */
  runs: TextRun[];
  /** 当前内容使用的渲染样式。 */
  style?: TextStyle;
  /** 段落的大纲层级。 */
  level?: number;
  /** 当前段落使用的项目符号配置。 */
  bullet?: TextBulletStyle;
};

/** 所有幻灯片绘制元素共享的几何和层叠属性。 */
export type BaseElement = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 用于区分联合类型分支的类型标识。 */
  type: string;
  /** 相对定位区域左侧的横坐标，单位为标准化渲染像素。 */
  x: number;
  /** 相对定位区域顶部的纵坐标，单位为标准化渲染像素。 */
  y: number;
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 顺时针旋转角度，单位为度。 */
  rotate?: number;
  /** 是否沿水平方向翻转对象。 */
  flipH?: boolean;
  /** 是否沿垂直方向翻转对象。 */
  flipV?: boolean;
  /** 层叠顺序，数值越大越靠近前景。 */
  zIndex?: number;
  /** 整体透明度，0 表示完全透明，1 表示完全不透明。 */
  opacity?: number;
  /** 元素继承的占位符类型。 */
  placeholderType?: string;
  /** 元素继承的占位符索引。 */
  placeholderIdx?: string;
};

/** 包含段落、形状外观和文本框样式的文字元素。 */
export type TextElement = BaseElement & {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'text';
  /** 按源文档顺序排列的段落。 */
  paragraphs: TextParagraph[];
  /** 文本框级别的文字和排版样式。 */
  boxStyle?: TextStyle;
  /** 文本框外形使用的预设几何名称。 */
  shape?: string;
  /** 在压缩包、复合文档或资源表中的路径。 */
  path?: string;
  /** 矢量路径使用的坐标范围。 */
  viewBox?: string;
  /** 填充颜色、渐变或无填充标记。 */
  fill?: string | GradientFill | null;
  /** 文本框填充透明度，取值范围为 0 到 1。 */
  fillOpacity?: number;
  /** 轮廓颜色；null 表示明确不绘制轮廓。 */
  stroke?: string | null;
  /** 文本框轮廓透明度，取值范围为 0 到 1。 */
  strokeOpacity?: number;
  /** 文本框轮廓宽度，单位为标准化渲染像素。 */
  strokeWidth?: number;
  /** 文本框轮廓使用的虚线样式。 */
  strokeDash?: string;
  /** 文本框的阴影样式。 */
  shadow?: ShadowStyle;
  /** 圆角半径；数值单位为标准化渲染像素。 */
  borderRadius?: number;
};

/** 包含几何路径、填充、轮廓和阴影的形状元素。 */
export type ShapeElement = BaseElement & {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'shape';
  /** 形状使用的预设几何名称。 */
  shape: string;
  /** 在压缩包、复合文档或资源表中的路径。 */
  path?: string;
  /** 矢量路径使用的坐标范围。 */
  viewBox?: string;
  /** 填充颜色、渐变或无填充标记。 */
  fill?: string | GradientFill | null;
  /** 形状填充透明度，取值范围为 0 到 1。 */
  fillOpacity?: number;
  /** 轮廓颜色；null 表示明确不绘制轮廓。 */
  stroke?: string | null;
  /** 形状轮廓透明度，取值范围为 0 到 1。 */
  strokeOpacity?: number;
  /** 形状轮廓宽度，单位为标准化渲染像素。 */
  strokeWidth?: number;
  /** 形状轮廓使用的虚线样式。 */
  strokeDash?: string;
  /** 形状的阴影样式。 */
  shadow?: ShadowStyle;
  /** 圆角半径；数值单位为标准化渲染像素。 */
  borderRadius?: number;
};

/** 包含图片资源、替代文本和裁剪范围的图片元素。 */
export type ImageElement = BaseElement & {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'image';
  /** 图片资源地址或延迟资源引用。 */
  src: string | OfficeResourceSource;
  /** 图片无法显示时使用的替代文本。 */
  alt?: string;
  /** 图片裁剪边界。 */
  crop?: ImageCrop;
};

/** 包含行列尺寸和单元格内容的表格元素。 */
export type TableElement = BaseElement & {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'table';
  /** 各表格列的宽度，单位为标准化渲染像素。 */
  columnWidths?: number[];
  /** 各表格行的测量高度。 */
  rowHeights?: number[];
  /** 按显示顺序排列的表格行。 */
  rows: TableCell[][];
};

/** 演示文稿表格单元格的内容和样式。 */
export type TableCell = {
  /** 文本内容。 */
  text: string;
  /** 按源文档顺序排列的段落。 */
  paragraphs?: TextParagraph[];
  /** 当前内容使用的渲染样式。 */
  style?: TextStyle;
  /** 背景颜色，使用 CSS 颜色值。 */
  backgroundColor?: string | null;
  /** 单元格背景透明度，取值范围为 0 到 1。 */
  backgroundOpacity?: number;
  /** 边框颜色。 */
  borderColor?: string | null;
  /** 单元格边框透明度，取值范围为 0 到 1。 */
  borderOpacity?: number;
  /** 边框宽度，单位为标准化渲染像素。 */
  borderWidth?: number;
  /** 四个方向的内边距配置。 */
  margins?: {
    /** 单元格左内边距，单位为标准化渲染像素。 */
    left?: number;
    /** 单元格右内边距，单位为标准化渲染像素。 */
    right?: number;
    /** 单元格上内边距，单位为标准化渲染像素。 */
    top?: number;
    /** 单元格下内边距，单位为标准化渲染像素。 */
    bottom?: number;
  };
  /** 垂直对齐方式。 */
  verticalAlign?: 'top' | 'middle' | 'bottom';
};

/** 共享变换和层叠关系的一组幻灯片元素。 */
export type GroupElement = BaseElement & {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'group';
  /** 按绘制顺序排列的子元素。 */
  children: SlideElement[];
};

/** 暂不支持还原但需保留占位信息的幻灯片元素。 */
export type UnsupportedElement = BaseElement & {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'unsupported';
  /** 暂不支持还原该元素的原因。 */
  reason: string;
};

/** 图片四个方向的裁剪范围。 */
export type ImageCrop = {
  /** 左侧位置或间距，单位由所属模型定义。 */
  left?: number;
  /** 顶部位置或间距，单位由所属模型定义。 */
  top?: number;
  /** 右侧位置或间距，单位由所属模型定义。 */
  right?: number;
  /** 底部位置或间距，单位由所属模型定义。 */
  bottom?: number;
};

/** 描述演示文稿标准模型使用的样式参数。 */
export type ShadowStyle = {
  /** 前景或文字颜色，使用 CSS 颜色值。 */
  color?: string;
  /** 整体透明度，0 表示完全透明，1 表示完全不透明。 */
  opacity?: number;
  /** 阴影模糊半径，单位为标准化渲染像素。 */
  blur?: number;
  /** 阴影的水平偏移，单位为标准化渲染像素。 */
  offsetX?: number;
  /** 阴影的垂直偏移，单位为标准化渲染像素。 */
  offsetY?: number;
};

/** 标准化幻灯片支持的绘制元素联合类型。 */
export type SlideElement =
  | TextElement
  | ShapeElement
  | ImageElement
  | ChartElement
  | TableElement
  | GroupElement
  | UnsupportedElement;

/** 包含图表模型和按需快照的图表元素。 */
export type ChartElement = BaseElement & {
  /** 用于区分联合类型分支的类型标识。 */
  type: 'chart';
  /** 图表渲染相关文案。 */
  chart: import('../../shared/ooxml/charts').OfficeChartModel;
  /** 图表在演示文稿中的稳定标识。 */
  chartId?: string;
  /** 图表定义在压缩包中的路径。 */
  chartPath?: string;
  /** 图表使用的按需静态快照资源。 */
  snapshotSource?: OfficeResourceSource;
};
import type { OfficeResourceSource } from '../resource-store';
