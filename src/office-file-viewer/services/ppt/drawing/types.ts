import type { TextStyle } from '../../presentation/types';

/** 描述 PptOfficeArtProperty 在 PPT 二进制解析中的数据结构。 */
export type PptOfficeArtProperty = {
  /** PptOfficeArtProperty 在所属文档或任务中的唯一标识。 */
  id: number;
  /** PptOfficeArtProperty 保存的解析值或业务值。 */
  value: number;
  /** OfficeArt 属性值是否引用图片资源。 */
  isBlip: boolean;
  /** PptOfficeArtProperty 的 complexData 原始二进制数据。 */
  complexData?: Uint8Array;
};

/** 描述 PPT 二进制解析使用的样式参数。 */
export type PptShapeStyle = {
  /** PptShapeStyle 的填充颜色、渐变或无填充标记；未提供时沿用来源格式或渲染器的默认规则。 */
  fill?: string | null;
  /** PptShapeStyle 的轮廓颜色；null 表示明确不绘制轮廓；未提供时沿用来源格式或渲染器的默认规则。 */
  stroke?: string | null;
  /** PptShapeStyle 的轮廓宽度，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  strokeWidth?: number;
  /** PptShapeStyle 的顺时针旋转角度，单位为度；未提供时沿用来源格式或渲染器的默认规则。 */
  rotate?: number;
  /** 是否沿水平方向翻转对象。 */
  flipH?: boolean;
  /** 是否沿垂直方向翻转对象。 */
  flipV?: boolean;
  /** PptShapeStyle 关联的 textStyle 结构；字段形状由 TextStyle 定义；未提供时使用来源格式或渲染器的默认行为。 */
  textStyle?: TextStyle;
};

/** 描述 PPT 二进制解析对象的位置和定位基准。 */
export type PptShapeAnchor = {
  /** PptShapeAnchor 的 x 几何值，单位遵循对应 Office 二进制记录定义。 */
  x: number;
  /** PptShapeAnchor 的 y 几何值，单位遵循对应 Office 二进制记录定义。 */
  y: number;
  /** PptShapeAnchor 的 width 几何值，单位遵循对应 Office 二进制记录定义。 */
  width: number;
  /** PptShapeAnchor 的 height 几何值，单位遵循对应 Office 二进制记录定义。 */
  height: number;
};
