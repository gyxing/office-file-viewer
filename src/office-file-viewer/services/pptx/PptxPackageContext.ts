import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import type { OfficeArchiveReader } from '../../shared/ooxml/OfficeArchiveReader';
import type {
  GradientFill,
  PresentationSlideDescriptor,
  ShadowStyle,
  SlideBackground,
  SlideElement,
  TextStyle,
  ThemeModel,
} from '../presentation/types';
import type { OfficeResourceSource } from '../resource-store';

/** PPTX 包内关系文件按路径组织的目标映射。 */
export type RelationshipMap = Record<string, Record<string, string>>;

/** PPTX 包解析期间共享且可按需补充的轻量状态。 */
export type PptxPackageState = {
  /** 按包内路径索引的 OOXML 条目。 */
  entries: OfficeEntryMap;
  /** 按关系文件路径组织的 OOXML 关系映射。 */
  relationships: RelationshipMap;
  /** 按媒体文件名索引的资源地址或延迟资源。 */
  mediaByName: Record<string, string | OfficeResourceSource>;
  /** 按包内完整路径索引的媒体资源。 */
  mediaByPath: Record<string, string | OfficeResourceSource>;
};

/** PPTX 表格样式支持的区域变体名称。 */
export type TableStyleVariantName =
  | 'wholeTbl'
  | 'band1H'
  | 'band2H'
  | 'band1V'
  | 'band2V'
  | 'firstRow'
  | 'lastRow'
  | 'firstCol'
  | 'lastCol';

/** PPTX 表格单元格的填充、边框和文字样式。 */
export type TableCellStyle = {
  /** 表格单元格内文字使用的样式。 */
  text?: TextStyle;
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
};

/** PPTX 表格样式及其区域变体。 */
export type TableStyleDefinition = {
  /** 样式标识。 */
  styleId: string;
  /** 来源文档中的表格样式名称。 */
  styleName?: string;
  /** 按表格区域名称索引的局部样式。 */
  variants: Partial<Record<TableStyleVariantName, TableCellStyle>>;
};

/** 按样式标识索引的 PPTX 表格样式。 */
export type TableStyleMap = Record<string, TableStyleDefinition>;

/** PPTX 占位符继承的文本和形状样式。 */
export type PlaceholderStyle = {
  /** 占位符类型，例如标题、正文或页码。 */
  type?: string;
  /** 版式内用于匹配占位符的标识。 */
  idx?: string;
  /** 相对定位区域左侧的横坐标，单位为标准化渲染像素。 */
  x?: number;
  /** 相对定位区域顶部的纵坐标，单位为标准化渲染像素。 */
  y?: number;
  /** 宽度，单位为标准化渲染像素。 */
  width?: number;
  /** 高度，单位为标准化渲染像素。 */
  height?: number;
  /** 填充颜色、渐变或无填充标记。 */
  fill?: string | GradientFill | null;
  /** 占位符填充透明度，取值范围为 0 到 1。 */
  fillOpacity?: number;
  /** 轮廓颜色；null 表示明确不绘制轮廓。 */
  stroke?: string | null;
  /** 占位符轮廓透明度，取值范围为 0 到 1。 */
  strokeOpacity?: number;
  /** 占位符轮廓宽度，单位为标准化渲染像素。 */
  strokeWidth?: number;
  /** 占位符轮廓的虚线样式。 */
  strokeDash?: string;
  /** 阴影。 */
  shadow?: ShadowStyle;
  /** 占位符自身定义的文字样式。 */
  text?: TextStyle;
  /** 占位符正文层级的默认文字样式。 */
  body?: TextStyle;
  /** 按大纲层级索引的文字样式。 */
  levels?: Record<number, TextStyle>;
};

/** PPTX 幻灯片版式及其占位符定义。 */
export type LayoutDefinition = {
  /** 在压缩包、复合文档或资源表中的路径。 */
  path: string;
  /** 母版路径。 */
  masterPath: string;
  /** 尚未加载内容所使用的占位项。 */
  placeholders: Record<string, PlaceholderStyle>;
  /** 按占位符用途索引的版式文字预设。 */
  textPresets: Record<string, PlaceholderStyle>;
  /** 当前页面、幻灯片或元素的背景配置。 */
  background?: SlideBackground;
  /** 按绘制顺序排列的演示文稿元素。 */
  elements: SlideElement[];
};

/** PPTX 幻灯片母版、版式和主题信息。 */
export type MasterDefinition = {
  /** 在压缩包、复合文档或资源表中的路径。 */
  path: string;
  /** 尚未加载内容所使用的占位项。 */
  placeholders: Record<string, PlaceholderStyle>;
  /** 按占位符用途索引的母版文字预设。 */
  textPresets: Record<string, PlaceholderStyle>;
  /** 当前页面、幻灯片或元素的背景配置。 */
  background?: SlideBackground;
  /** 按绘制顺序排列的演示文稿元素。 */
  elements: SlideElement[];
};

/** 在公共描述符之外保留 PPTX 单页读取所需的包内路径。 */
export type PptxSlideDescriptor = PresentationSlideDescriptor & {
  /** 幻灯片路径。 */
  slidePath: string;
  /** 当前幻灯片关系文件的包内路径。 */
  relsPath: string;
  /** 备注路径。 */
  notesPath?: string;
};

/** 大型 PPTX 生命周期内常驻的结构、主题和包 Reader。 */
export type PptxPackageContext = {
  /** 当前解析或预览会话的标识。 */
  sessionId: string;
  /** 用于按需读取源数据的读取器。 */
  reader: OfficeArchiveReader;
  /** 解析 PPTX 各部件时共享的包状态。 */
  packageState: PptxPackageState;
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 当前文档使用的主题颜色和字体配置。 */
  theme: ThemeModel;
  /** 按样式标识索引的表格样式。 */
  tableStyles: TableStyleMap;
  /** 可供幻灯片继承的母版定义。 */
  masterDefinitions: readonly MasterDefinition[];
  /** 可供幻灯片继承的版式定义。 */
  layoutDefinitions: readonly LayoutDefinition[];
  /** 按源顺序排列的轻量描述信息。 */
  descriptors: readonly PptxSlideDescriptor[];
};
