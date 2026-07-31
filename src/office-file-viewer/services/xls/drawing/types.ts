import type { SpreadsheetWarning } from '../../spreadsheet/types';
export type { OfficeArtRecord } from '../../../shared/officeart';

/** BIFF8 浮动对象锚点的单元格和偏移比例。 */
export type Biff8AnchorPoint = {
  /** 锚点所在的零基行索引。 */
  row: number;
  /** 锚点所在的零基列索引。 */
  column: number;
  /** 在锚点单元格内的相对位置比例。 */
  rowFraction: number;
  /** 在锚点单元格内的相对位置比例。 */
  columnFraction: number;
};

/** BIFF8 浮动对象的起止锚点。 */
export type Biff8Anchor = {
  /** 起始锚点。 */
  from: Biff8AnchorPoint;
  /** 结束锚点。 */
  to: Biff8AnchorPoint;
};

/** BIFF8 绘图图片支持的编码格式。 */
export type Biff8DrawingImageFormat =
  | 'png'
  | 'jpeg'
  | 'gif'
  | 'dib'
  | 'wmf'
  | 'emf'
  | 'pict'
  | 'unknown';

/** BIFF8 绘图图片资源及其定位信息。 */
export type Biff8DrawingImage = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 面向用户展示的名称。 */
  name?: string;
  /** 图片内容的实际编码格式。 */
  format: Biff8DrawingImageFormat;
  /** 图片或图元文件的原始字节。 */
  bytes: Uint8Array;
  /** 对象在工作表或画布中的定位锚点。 */
  anchor: Biff8Anchor;
  /** 图片无法显示时使用的替代文本。 */
  alt?: string;
  /** 源负载是否使用压缩存储。 */
  compressed?: boolean;
  /** 解析时产生但不阻止继续预览的警告。 */
  warnings: SpreadsheetWarning[];
};

/** BIFF8 OfficeArt 形状及其锚点和外观。 */
export type Biff8DrawingShape = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 绘图形状在 OfficeArt 数据中的标识。 */
  shapeId?: number;
  /** OfficeArt 自动形状类型编号。 */
  shapeType?: number;
  /** 面向用户展示的名称。 */
  name?: string;
  /** 形状引用的 OfficeArt 图片索引。 */
  blipIndex?: number;
  /** OfficeArt 中以 BGR 顺序编码的形状填充色。 */
  fillColor?: number;
  /** OfficeArt 中以 BGR 顺序编码的形状轮廓色。 */
  lineColor?: number;
  /** 形状轮廓宽度，单位为 EMU。 */
  lineWidth?: number;
  /** 对象在工作表或画布中的定位锚点。 */
  anchor: Biff8Anchor;
};

/** 完成 DIB 解码的 RGBA 位图。 */
export type DecodedBitmap = {
  /** 解码后位图的像素宽度。 */
  width: number;
  /** 解码后位图的像素高度。 */
  height: number;
  /** 按 RGBA 顺序存储的解码像素字节。 */
  rgba: Uint8ClampedArray;
};

/** 从 OfficeArt 图片存储解析出的图片。 */
export type ParsedBlip = {
  /** 图片在 OfficeArt 图片存储中的索引。 */
  index: number;
  /** 图片内容的实际编码格式。 */
  format: Biff8DrawingImageFormat;
  /** 图片或图元文件的原始字节。 */
  bytes: Uint8Array;
  /** 源负载是否使用压缩存储。 */
  compressed?: boolean;
  /** 解析时产生但不阻止继续预览的警告。 */
  warnings: SpreadsheetWarning[];
};
