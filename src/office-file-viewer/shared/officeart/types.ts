/** 描述 OfficeArt 图形过程中可继续处理的警告。 */
export type OfficeArtWarning = {
  /** OfficeArtWarning 的稳定代码，用于程序化识别具体情况。 */
  code: string;
  /** OfficeArtWarning 面向调用方或用户展示的具体警告、错误说明。 */
  message: string;
  /** OfficeArtWarning 在源二进制流中的字节偏移。 */
  offset?: number;
};

/** 表示OfficeArt 图形读取到的一条记录。 */
export type OfficeArtRecord = {
  /** 消息或数据结构采用的协议版本号。 */
  version: number;
  /** OfficeArtRecord 从源格式读取的 instance 枚举或标识值。 */
  instance: number;
  /** 用于区分 OfficeArtRecord 不同结构分支的类型标识。 */
  type: number;
  /** OfficeArtRecord 对应二进制记录或数据块的字节长度。 */
  length: number;
  /** OfficeArtRecord 在源二进制流中的字节偏移。 */
  offset: number;
  /** OfficeArtRecord 当前步骤需要处理的原始或标准化数据。 */
  data: Uint8Array;
  /** OfficeArtRecord 包含并负责布局的 React 子节点；未提供时使用来源格式或渲染器的默认行为。 */
  children?: OfficeArtRecord[];
};

/** 描述 OfficeArtImageFormat 在OfficeArt 图形中的数据结构。 */
export type OfficeArtImageFormat =
  | 'png'
  | 'jpeg'
  | 'gif'
  | 'dib'
  | 'wmf'
  | 'emf'
  | 'pict'
  | 'unknown';

/** 描述 DecodedBitmap 在OfficeArt 图形中的数据结构。 */
export type DecodedBitmap = {
  /** DecodedBitmap 的 width 几何值，单位遵循对应 Office 二进制记录定义。 */
  width: number;
  /** DecodedBitmap 的 height 几何值，单位遵循对应 Office 二进制记录定义。 */
  height: number;
  /** DecodedBitmap 按 RGBA 顺序存储的解码像素字节。 */
  rgba: Uint8ClampedArray;
};

/** 描述 ParsedOfficeArtBlip 在OfficeArt 图形中的数据结构。 */
export type ParsedOfficeArtBlip = {
  /** ParsedOfficeArtBlip 在所属集合中的位置索引。 */
  index: number;
  /** ParsedOfficeArtBlip 的文件格式或协议类型标识。 */
  format: OfficeArtImageFormat;
  /** ParsedOfficeArtBlip 保存的原始字节序列。 */
  bytes: Uint8Array;
  /** 表示 ParsedOfficeArtBlip 的负载是否采用压缩存储。 */
  compressed?: boolean;
  /** ParsedOfficeArtBlip 解析时产生但不阻止继续预览的警告集合。 */
  warnings: OfficeArtWarning[];
};
