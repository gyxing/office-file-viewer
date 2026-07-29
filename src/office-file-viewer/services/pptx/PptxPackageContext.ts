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
  entries: OfficeEntryMap;
  relationships: RelationshipMap;
  mediaByName: Record<string, string | OfficeResourceSource>;
  mediaByPath: Record<string, string | OfficeResourceSource>;
};

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

export type TableCellStyle = {
  text?: TextStyle;
  backgroundColor?: string | null;
  backgroundOpacity?: number;
  borderColor?: string | null;
  borderOpacity?: number;
  borderWidth?: number;
};

export type TableStyleDefinition = {
  styleId: string;
  styleName?: string;
  variants: Partial<Record<TableStyleVariantName, TableCellStyle>>;
};

export type TableStyleMap = Record<string, TableStyleDefinition>;

export type PlaceholderStyle = {
  type?: string;
  idx?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  fill?: string | GradientFill | null;
  fillOpacity?: number;
  stroke?: string | null;
  strokeOpacity?: number;
  strokeWidth?: number;
  strokeDash?: string;
  shadow?: ShadowStyle;
  text?: TextStyle;
  body?: TextStyle;
  levels?: Record<number, TextStyle>;
};

export type LayoutDefinition = {
  path: string;
  masterPath: string;
  placeholders: Record<string, PlaceholderStyle>;
  textPresets: Record<string, PlaceholderStyle>;
  background?: SlideBackground;
  elements: SlideElement[];
};

export type MasterDefinition = {
  path: string;
  placeholders: Record<string, PlaceholderStyle>;
  textPresets: Record<string, PlaceholderStyle>;
  background?: SlideBackground;
  elements: SlideElement[];
};

/** 在公共描述符之外保留 PPTX 单页读取所需的包内路径。 */
export type PptxSlideDescriptor = PresentationSlideDescriptor & {
  slidePath: string;
  relsPath: string;
  notesPath?: string;
};

/** 大型 PPTX 生命周期内常驻的结构、主题和包 Reader。 */
export type PptxPackageContext = {
  sessionId: string;
  reader: OfficeArchiveReader;
  packageState: PptxPackageState;
  width: number;
  height: number;
  theme: ThemeModel;
  tableStyles: TableStyleMap;
  masterDefinitions: readonly MasterDefinition[];
  layoutDefinitions: readonly LayoutDefinition[];
  descriptors: readonly PptxSlideDescriptor[];
};
