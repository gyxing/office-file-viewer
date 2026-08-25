/** 文档内容区水印配置。 */
export type OfficeViewerWatermarkOptions = Readonly<{
  /** 水印文字；数组用于渲染多行内容。 */
  content: string | readonly string[];
  /** 水印文字颜色，默认使用中性灰。 */
  color?: string;
  /** 水印透明度，取值范围为 0 至 1。 */
  opacity?: number;
  /** 水印字号，单位为像素。 */
  fontSize?: number;
  /** 水印字体族。 */
  fontFamily?: string;
  /** 水印字重。 */
  fontWeight?: string | number;
  /** 水印顺时针旋转角度。 */
  rotate?: number;
  /** 相邻水印内容的水平、垂直间距。 */
  gap?: readonly [number, number];
  /** 水印图案相对默认位置的水平、垂直偏移。 */
  offset?: readonly [number, number];
}>;

/** 经过边界约束后供渲染层直接消费的水印配置。 */
export type ResolvedOfficeViewerWatermark = Readonly<{
  content: readonly string[];
  color: string;
  opacity: number;
  fontSize: number;
  fontFamily: string;
  fontWeight: string | number;
  rotate: number;
  gap: readonly [number, number];
  offset: readonly [number, number];
}>;

/** 允许关闭或配置水印的公共属性值。 */
export type OfficeViewerWatermark = false | OfficeViewerWatermarkOptions;
