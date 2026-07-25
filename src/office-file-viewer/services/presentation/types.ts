/** 描述演示文稿标准模型过程中可继续处理的警告。 */
export type PresentationWarning = {
  /** PresentationWarning 的稳定代码，用于程序化识别具体情况。 */
  code: string;
  /** PresentationWarning 面向调用方或用户展示的具体警告、错误说明。 */
  message: string;
  /** 幻灯片在演示文稿集合中的索引。 */
  slideIndex?: number;
  /** PresentationWarning 的 objectId 文本值。 */
  objectId?: string;
  /** PresentationWarning 在所属数据范围中的偏移位置。 */
  offset?: number;
};

/** 记录演示文稿标准模型持有且需要统一管理的资源。 */
export type PresentationResources = {
  /** 浏览器创建的对象 URL 集合，文档释放时必须逐一撤销。 */
  objectUrls: string[];
};

/** 描述演示文稿标准模型生成的标准化文档模型。 */
export type PresentationDocument = {
  /** PresentationDocument 的 width 尺寸或坐标，单位为标准化渲染像素。 */
  width: number;
  /** PresentationDocument 的 height 尺寸或坐标，单位为标准化渲染像素。 */
  height: number;
  /** PresentationDocument 使用的主题颜色和字体配置。 */
  theme: ThemeModel;
  /** PresentationDocument 包含的 slides 有序集合。 */
  slides: SlideModel[];
  /** 解析时产生但不阻止继续预览的警告；未提供表示没有警告。 */
  warnings?: PresentationWarning[];
  /** 文档持有且需要在销毁时释放的浏览器资源；未提供表示没有需释放资源。 */
  resources?: PresentationResources;
};

/** 描述演示文稿标准模型使用的标准化模型。 */
export type ThemeModel = {
  /** 按主题槽名称索引的标准化颜色映射。 */
  colorScheme: Record<string, string>;
  /** 按主题字体槽名称索引的字体族映射。 */
  fontScheme: Record<string, string>;
  /** 将幻灯片颜色槽重定向到主题颜色槽的映射。 */
  colorMap?: Record<string, string>;
};

/** 描述 GradientStop 在演示文稿标准模型中的数据结构。 */
export type GradientStop = {
  /** 渐变停止点的位置比例，取值范围为 0 到 1。 */
  offset: number;
  /** GradientStop 的前景或文本颜色，使用标准化 CSS 颜色值。 */
  color: string;
};

/** 描述 GradientFill 在演示文稿标准模型中的数据结构。 */
export type GradientFill = {
  /** 用于区分 GradientFill 不同结构分支的类型标识。 */
  type: 'linear';
  /** GradientFill 的角度值，单位为度。 */
  angle: number;
  /** GradientFill 包含的 stops 有序集合。 */
  stops: GradientStop[];
};

/** 描述演示文稿标准模型使用的标准化模型。 */
export type SlideModel = {
  /** SlideModel 在所属文档或任务中的唯一标识。 */
  id: string;
  /** 幻灯片在演示文稿中的稳定顺序索引。 */
  index: number;
  /** SlideModel 的 width 尺寸或坐标，单位为标准化渲染像素。 */
  width: number;
  /** SlideModel 的 height 尺寸或坐标，单位为标准化渲染像素。 */
  height: number;
  /** 是否隐藏 SlideModel；未提供时沿用来源格式或渲染器的默认规则。 */
  hidden?: boolean;
  /** SlideModel 的背景填充模型；未提供时使用来源格式或渲染器的默认行为。 */
  background?: SlideBackground;
  /** SlideModel 包含的 elements 有序集合。 */
  elements: SlideElement[];
};

/** 描述 SlideBackground 在演示文稿标准模型中的数据结构。 */
export type SlideBackground = {
  /** SlideBackground 的填充颜色、渐变或无填充标记；未提供时沿用来源格式或渲染器的默认规则。 */
  fill?: string;
  /** SlideBackground 填充区域的透明度，取值范围为 0 到 1；未提供时沿用来源格式或渲染器的默认规则。 */
  fillOpacity?: number;
  /** 用作幻灯片背景的媒体资源引用。 */
  imageRef?: string;
};

/** 描述演示文稿标准模型使用的样式参数。 */
export type TextStyle = {
  /** TextStyle 的字体族名称；未提供时沿用来源格式或渲染器的默认规则。 */
  fontFamily?: string;
  /** TextStyle 的字号，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  fontSize?: number;
  /** 是否使用粗体渲染 TextStyle；未提供时沿用来源格式或渲染器的默认规则。 */
  bold?: boolean;
  /** 是否使用斜体渲染 TextStyle；未提供时沿用来源格式或渲染器的默认规则。 */
  italic?: boolean;
  /** 是否为 TextStyle 绘制下划线；未提供时沿用来源格式或渲染器的默认规则。 */
  underline?: boolean;
  /** TextStyle 使用的删除线类型；未提供时沿用来源格式或渲染器的默认规则。 */
  strike?: 'none' | 'sngStrike' | 'dblStrike';
  /** 是否将小写字母以小型大写字形显示。 */
  smallCaps?: boolean;
  /** 是否将文本全部以大写字形显示。 */
  allCaps?: boolean;
  /** TextStyle 的前景或文本颜色，使用标准化 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
  color?: string;
  /** TextStyle 的整体透明度，0 表示完全透明，1 表示完全不透明；未提供时沿用来源格式或渲染器的默认规则。 */
  opacity?: number;
  /** TextStyle 的水平对齐方式；未提供时沿用来源格式或渲染器的默认规则。 */
  align?: 'left' | 'center' | 'right' | 'justify';
  /** TextStyle 的垂直对齐方式；未提供时沿用来源格式或渲染器的默认规则。 */
  verticalAlign?: 'top' | 'middle' | 'bottom';
  /** 文本书写方向；未提供时采用水平从上到下。 */
  writingMode?: 'horizontal-tb' | 'vertical-rl' | 'vertical-lr';
  /** 文本溢出时采用不调整、缩小文本或扩大形状的策略。 */
  fit?: 'none' | 'shrinkText' | 'resizeShape';
  /** TextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  marginLeft?: number;
  /** TextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  marginRight?: number;
  /** TextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  marginTop?: number;
  /** TextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  marginBottom?: number;
  /** TextStyle 的行高，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  lineHeight?: number;
  /** TextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  spaceBefore?: number;
  /** TextStyle 的对应间距，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  spaceAfter?: number;
  /** 段落首行文本缩进，单位为标准化渲染像素。 */
  textIndent?: number;
  /** TextStyle 的 charSpace 渲染数值，单位为标准化渲染像素。 */
  charSpace?: number;
  /** TextStyle 的 baseline 渲染数值，单位为标准化渲染像素。 */
  baseline?: number;
  /** 段落项目符号样式；未提供表示不覆盖段落级设置。 */
  bullet?: TextBulletStyle;
};

/** 描述演示文稿标准模型使用的样式参数。 */
export type TextBulletStyle = {
  /** TextBulletStyle 的 char 文本值。 */
  char?: string;
  /** TextBulletStyle 的前景或文本颜色，使用标准化 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
  color?: string;
  /** 项目符号字号，单位为标准化渲染像素。 */
  size?: number;
  /** 是否显式禁用项目符号。 */
  none?: boolean;
};

/** 描述 TextRun 在演示文稿标准模型中的数据结构。 */
export type TextRun = {
  /** TextRun 携带或渲染的文本内容。 */
  text: string;
  /** TextRun 使用的渲染或文本样式。 */
  style?: TextStyle;
};

/** 描述 TextParagraph 在演示文稿标准模型中的数据结构。 */
export type TextParagraph = {
  /** TextParagraph 包含的 runs 有序集合。 */
  runs: TextRun[];
  /** TextParagraph 使用的渲染或文本样式。 */
  style?: TextStyle;
  /** 段落的大纲层级。 */
  level?: number;
  /** 当前段落使用的项目符号配置。 */
  bullet?: TextBulletStyle;
};

/** 描述 BaseElement 在演示文稿标准模型中的数据结构。 */
export type BaseElement = {
  /** BaseElement 在所属文档或任务中的唯一标识。 */
  id: string;
  /** 用于区分 BaseElement 不同结构分支的类型标识。 */
  type: string;
  /** BaseElement 的 x 尺寸或坐标，单位为标准化渲染像素。 */
  x: number;
  /** BaseElement 的 y 尺寸或坐标，单位为标准化渲染像素。 */
  y: number;
  /** BaseElement 的 width 尺寸或坐标，单位为标准化渲染像素。 */
  width: number;
  /** BaseElement 的 height 尺寸或坐标，单位为标准化渲染像素。 */
  height: number;
  /** BaseElement 的顺时针旋转角度，单位为度；未提供时沿用来源格式或渲染器的默认规则。 */
  rotate?: number;
  /** 是否沿水平方向翻转对象。 */
  flipH?: boolean;
  /** 是否沿垂直方向翻转对象。 */
  flipV?: boolean;
  /** BaseElement 的层叠顺序，数值越大越靠近前景。 */
  zIndex?: number;
  /** BaseElement 的整体透明度，0 表示完全透明，1 表示完全不透明；未提供时沿用来源格式或渲染器的默认规则。 */
  opacity?: number;
  /** 元素继承的占位符类型。 */
  placeholderType?: string;
  /** 元素继承的占位符索引。 */
  placeholderIdx?: string;
};

/** 描述 TextElement 在演示文稿标准模型中的数据结构。 */
export type TextElement = BaseElement & {
  /** 用于区分 TextElement 不同结构分支的类型标识。 */
  type: 'text';
  /** TextElement 包含的 paragraphs 有序集合。 */
  paragraphs: TextParagraph[];
  /** TextElement 关联的 boxStyle 结构；字段形状由 TextStyle 定义；未提供时使用来源格式或渲染器的默认行为。 */
  boxStyle?: TextStyle;
  /** TextElement 的 shape 文本值。 */
  shape?: string;
  /** TextElement 在压缩包、复合文档或图形数据中的路径。 */
  path?: string;
  /** TextElement 的 viewBox 文本值。 */
  viewBox?: string;
  /** TextElement 的填充颜色、渐变或无填充标记；未提供时沿用来源格式或渲染器的默认规则。 */
  fill?: string | GradientFill | null;
  /** TextElement 填充区域的透明度，取值范围为 0 到 1；未提供时沿用来源格式或渲染器的默认规则。 */
  fillOpacity?: number;
  /** TextElement 的轮廓颜色；null 表示明确不绘制轮廓；未提供时沿用来源格式或渲染器的默认规则。 */
  stroke?: string | null;
  /** TextElement 轮廓的透明度，取值范围为 0 到 1；未提供时沿用来源格式或渲染器的默认规则。 */
  strokeOpacity?: number;
  /** TextElement 的轮廓宽度，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  strokeWidth?: number;
  /** TextElement 的轮廓虚线样式；未提供时沿用来源格式或渲染器的默认规则。 */
  strokeDash?: string;
  /** TextElement 的阴影样式；未提供时沿用来源格式或渲染器的默认规则。 */
  shadow?: ShadowStyle;
  /** TextElement 的 borderRadius 渲染尺寸，单位为标准化像素；未提供时使用来源格式或渲染器的默认行为。 */
  borderRadius?: number;
};

/** 描述 ShapeElement 在演示文稿标准模型中的数据结构。 */
export type ShapeElement = BaseElement & {
  /** 用于区分 ShapeElement 不同结构分支的类型标识。 */
  type: 'shape';
  /** ShapeElement 的 shape 文本值。 */
  shape: string;
  /** ShapeElement 在压缩包、复合文档或图形数据中的路径。 */
  path?: string;
  /** ShapeElement 的 viewBox 文本值。 */
  viewBox?: string;
  /** ShapeElement 的填充颜色、渐变或无填充标记；未提供时沿用来源格式或渲染器的默认规则。 */
  fill?: string | GradientFill | null;
  /** ShapeElement 填充区域的透明度，取值范围为 0 到 1；未提供时沿用来源格式或渲染器的默认规则。 */
  fillOpacity?: number;
  /** ShapeElement 的轮廓颜色；null 表示明确不绘制轮廓；未提供时沿用来源格式或渲染器的默认规则。 */
  stroke?: string | null;
  /** ShapeElement 轮廓的透明度，取值范围为 0 到 1；未提供时沿用来源格式或渲染器的默认规则。 */
  strokeOpacity?: number;
  /** ShapeElement 的轮廓宽度，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  strokeWidth?: number;
  /** ShapeElement 的轮廓虚线样式；未提供时沿用来源格式或渲染器的默认规则。 */
  strokeDash?: string;
  /** ShapeElement 的阴影样式；未提供时沿用来源格式或渲染器的默认规则。 */
  shadow?: ShadowStyle;
  /** ShapeElement 的 borderRadius 渲染尺寸，单位为标准化像素；未提供时使用来源格式或渲染器的默认行为。 */
  borderRadius?: number;
};

/** 描述 ImageElement 在演示文稿标准模型中的数据结构。 */
export type ImageElement = BaseElement & {
  /** 用于区分 ImageElement 不同结构分支的类型标识。 */
  type: 'image';
  /** ImageElement 的 src 文本值。 */
  src: string;
  /** ImageElement 的 alt 文本值。 */
  alt?: string;
  /** ImageElement 的图片裁剪边界；未提供时使用来源格式或渲染器的默认行为。 */
  crop?: ImageCrop;
};

/** 描述 TableElement 在演示文稿标准模型中的数据结构。 */
export type TableElement = BaseElement & {
  /** 用于区分 TableElement 不同结构分支的类型标识。 */
  type: 'table';
  /** TableElement 包含的 columnWidths 有序集合。 */
  columnWidths?: number[];
  /** TableElement 包含的 rowHeights 有序集合。 */
  rowHeights?: number[];
  /** TableElement 包含的 rows 有序集合。 */
  rows: TableCell[][];
};

/** 描述 TableCell 在演示文稿标准模型中的数据结构。 */
export type TableCell = {
  /** TableCell 携带或渲染的文本内容。 */
  text: string;
  /** TableCell 包含的 paragraphs 有序集合。 */
  paragraphs?: TextParagraph[];
  /** TableCell 使用的渲染或文本样式。 */
  style?: TextStyle;
  /** TableCell 的背景颜色，使用 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
  backgroundColor?: string | null;
  /** TableCell 的透明度，取值范围为 0 到 1。 */
  backgroundOpacity?: number;
  /** TableCell 的边框颜色配置；未提供时使用来源格式或渲染器的默认行为。 */
  borderColor?: string | null;
  /** TableCell 的透明度，取值范围为 0 到 1。 */
  borderOpacity?: number;
  /** TableCell 的 borderWidth 渲染数值，单位为标准化渲染像素。 */
  borderWidth?: number;
  /** TableCell 四个方向的内边距配置；未提供时使用来源格式或渲染器的默认行为。 */
  margins?: {
    /** TableCell 的 left 尺寸或坐标，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
    left?: number;
    /** TableCell 的 right 尺寸或坐标，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
    right?: number;
    /** TableCell 的 top 尺寸或坐标，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
    top?: number;
    /** TableCell 的 bottom 尺寸或坐标，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
    bottom?: number;
  };
  /** TableCell 的垂直对齐方式；未提供时沿用来源格式或渲染器的默认规则。 */
  verticalAlign?: 'top' | 'middle' | 'bottom';
};

/** 描述 GroupElement 在演示文稿标准模型中的数据结构。 */
export type GroupElement = BaseElement & {
  /** 用于区分 GroupElement 不同结构分支的类型标识。 */
  type: 'group';
  /** GroupElement 包含并负责布局的 React 子节点。 */
  children: SlideElement[];
};

/** 描述 UnsupportedElement 在演示文稿标准模型中的数据结构。 */
export type UnsupportedElement = BaseElement & {
  /** 用于区分 UnsupportedElement 不同结构分支的类型标识。 */
  type: 'unsupported';
  /** UnsupportedElement 的 reason 文本值。 */
  reason: string;
};

/** 描述 ImageCrop 在演示文稿标准模型中的数据结构。 */
export type ImageCrop = {
  /** ImageCrop 的 left 尺寸或坐标，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  left?: number;
  /** ImageCrop 的 top 尺寸或坐标，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  top?: number;
  /** ImageCrop 的 right 尺寸或坐标，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  right?: number;
  /** ImageCrop 的 bottom 尺寸或坐标，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  bottom?: number;
};

/** 描述演示文稿标准模型使用的样式参数。 */
export type ShadowStyle = {
  /** ShadowStyle 的前景或文本颜色，使用标准化 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
  color?: string;
  /** ShadowStyle 的整体透明度，0 表示完全透明，1 表示完全不透明；未提供时沿用来源格式或渲染器的默认规则。 */
  opacity?: number;
  /** ShadowStyle 的 blur 渲染尺寸，单位为标准化像素；未提供时使用来源格式或渲染器的默认行为。 */
  blur?: number;
  /** ShadowStyle 在对应二进制流中的字节偏移；未提供时使用来源格式或渲染器的默认行为。 */
  offsetX?: number;
  /** ShadowStyle 在对应二进制流中的字节偏移；未提供时使用来源格式或渲染器的默认行为。 */
  offsetY?: number;
};

/** 描述 SlideElement 在演示文稿标准模型中的数据结构。 */
export type SlideElement =
  | TextElement
  | ShapeElement
  | ImageElement
  | ChartElement
  | TableElement
  | GroupElement
  | UnsupportedElement;

/** 描述 ChartElement 在演示文稿标准模型中的数据结构。 */
export type ChartElement = BaseElement & {
  /** 用于区分 ChartElement 不同结构分支的类型标识。 */
  type: 'chart';
  /** ChartElement 当前关联的图表模型。 */
  chart: import('../../shared/ooxml/charts').OfficeChartModel;
  /** ChartElement 的 chartId 文本值。 */
  chartId?: string;
  /** ChartElement 的 chartPath 文本值。 */
  chartPath?: string;
};
