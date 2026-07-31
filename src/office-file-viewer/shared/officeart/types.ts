/** OfficeArt 解析过程中产生的非阻断警告。 */
export type OfficeArtWarning = {
  /** 供程序识别当前情况的稳定代码。 */
  code: string;
  /** 面向调用方或用户展示的说明。 */
  message: string;
  /** 在所属数据范围中的偏移位置。 */
  offset?: number;
};

/** OfficeArt 二进制流中的单条记录。 */
export type OfficeArtRecord = {
  /** 消息或数据结构采用的协议版本号。 */
  version: number;
  /** 当前二进制记录的实例字段。 */
  instance: number;
  /** OfficeArt 二进制记录的类型编号。 */
  type: number;
  /** 当前数据的长度。 */
  length: number;
  /** 在所属数据范围中的偏移位置。 */
  offset: number;
  /** 当前 OfficeArt 记录正文的原始字节。 */
  data: Uint8Array;
  /** 容器记录按源顺序包含的子记录。 */
  children?: OfficeArtRecord[];
};

/** OfficeArt 图片存储支持的编码格式。 */
export type OfficeArtImageFormat =
  | 'png'
  | 'jpeg'
  | 'gif'
  | 'dib'
  | 'wmf'
  | 'emf'
  | 'pict'
  | 'unknown';

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
export type ParsedOfficeArtBlip = {
  /** 图片在 OfficeArt 图片存储中的索引。 */
  index: number;
  /** 图片内容的实际编码格式。 */
  format: OfficeArtImageFormat;
  /** 图片或图元文件的原始字节。 */
  bytes: Uint8Array;
  /** 源负载是否使用压缩存储。 */
  compressed?: boolean;
  /** 解析时产生但不阻止继续预览的警告。 */
  warnings: OfficeArtWarning[];
};
