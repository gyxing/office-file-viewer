import type { SpreadsheetWarning } from '../../spreadsheet/types';
export type { OfficeArtRecord } from '../../../shared/officeart';

/** 描述 Biff8AnchorPoint 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8AnchorPoint = {
  /** Biff8AnchorPoint 使用的零基行列索引。 */
  row: number;
  /** Biff8AnchorPoint 使用的零基行列索引。 */
  column: number;
  /** Biff8AnchorPoint 在锚点单元格内的相对位置比例。 */
  rowFraction: number;
  /** Biff8AnchorPoint 在锚点单元格内的相对位置比例。 */
  columnFraction: number;
};

/** 描述 XLS/BIFF8 解析对象的位置和定位基准。 */
export type Biff8Anchor = {
  /** Biff8Anchor 的起始锚点。 */
  from: Biff8AnchorPoint;
  /** Biff8Anchor 的结束锚点。 */
  to: Biff8AnchorPoint;
};

/** 描述 Biff8DrawingImageFormat 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8DrawingImageFormat =
  | 'png'
  | 'jpeg'
  | 'gif'
  | 'dib'
  | 'wmf'
  | 'emf'
  | 'pict'
  | 'unknown';

/** 描述 Biff8DrawingImage 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8DrawingImage = {
  /** Biff8DrawingImage 在所属文档或任务中的唯一标识。 */
  id: string;
  /** Biff8DrawingImage 的可读名称。 */
  name?: string;
  /** Biff8DrawingImage 的文件格式或协议类型标识。 */
  format: Biff8DrawingImageFormat;
  /** Biff8DrawingImage 保存的原始字节序列。 */
  bytes: Uint8Array;
  /** Biff8DrawingImage 在工作表或画布中的定位锚点。 */
  anchor: Biff8Anchor;
  /** Biff8DrawingImage 的 alt 文本值。 */
  alt?: string;
  /** 表示 Biff8DrawingImage 的负载是否采用压缩存储。 */
  compressed?: boolean;
  /** Biff8DrawingImage 解析时产生但不阻止继续预览的警告集合。 */
  warnings: SpreadsheetWarning[];
};

/** 描述 Biff8DrawingShape 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8DrawingShape = {
  /** Biff8DrawingShape 在所属文档或任务中的唯一标识。 */
  id: string;
  /** Biff8DrawingShape 在源文件记录中的数字标识。 */
  shapeId?: number;
  /** Biff8DrawingShape 从源格式读取的 shapeType 枚举或标识值。 */
  shapeType?: number;
  /** Biff8DrawingShape 的可读名称。 */
  name?: string;
  /** Biff8DrawingShape 在所属集合中的位置索引。 */
  blipIndex?: number;
  /** Biff8DrawingShape 在工作表或画布中的定位锚点。 */
  anchor: Biff8Anchor;
};

/** 描述 DecodedBitmap 在 XLS/BIFF8 解析中的数据结构。 */
export type DecodedBitmap = {
  /** DecodedBitmap 的 width 几何值，单位遵循对应 Office 二进制记录定义。 */
  width: number;
  /** DecodedBitmap 的 height 几何值，单位遵循对应 Office 二进制记录定义。 */
  height: number;
  /** DecodedBitmap 按 RGBA 顺序存储的解码像素字节。 */
  rgba: Uint8ClampedArray;
};

/** 描述 ParsedBlip 在 XLS/BIFF8 解析中的数据结构。 */
export type ParsedBlip = {
  /** ParsedBlip 在所属集合中的位置索引。 */
  index: number;
  /** ParsedBlip 的文件格式或协议类型标识。 */
  format: Biff8DrawingImageFormat;
  /** ParsedBlip 保存的原始字节序列。 */
  bytes: Uint8Array;
  /** 表示 ParsedBlip 的负载是否采用压缩存储。 */
  compressed?: boolean;
  /** ParsedBlip 解析时产生但不阻止继续预览的警告集合。 */
  warnings: SpreadsheetWarning[];
};
