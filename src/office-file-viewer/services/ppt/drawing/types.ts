import type { TextStyle } from '../../presentation/types';

/** PPT OfficeArt 属性编号、数值和复杂数据。 */
export type PptOfficeArtProperty = {
  /** 在所属集合中的唯一标识。 */
  id: number;
  /** OfficeArt 属性的标量值。 */
  value: number;
  /** OfficeArt 属性值是否引用图片资源。 */
  isBlip: boolean;
  /** OfficeArt 属性携带的可选复杂字节数据。 */
  complexData?: Uint8Array;
};

/** PPT 形状的填充、轮廓和文字框样式。 */
export type PptShapeStyle = {
  /** 填充颜色、渐变或无填充标记。 */
  fill?: string | null;
  /** 轮廓颜色；null 表示明确不绘制轮廓。 */
  stroke?: string | null;
  /** 轮廓宽度，单位为标准化渲染像素。 */
  strokeWidth?: number;
  /** 顺时针旋转角度，单位为度。 */
  rotate?: number;
  /** 是否沿水平方向翻转对象。 */
  flipH?: boolean;
  /** 是否沿垂直方向翻转对象。 */
  flipV?: boolean;
  /** 形状内文字使用的基础样式。 */
  textStyle?: TextStyle;
};

/** PPT 形状在幻灯片坐标系中的位置和尺寸。 */
export type PptShapeAnchor = {
  /** 相对定位区域左侧的横坐标，单位为标准化渲染像素。 */
  x: number;
  /** 相对定位区域顶部的纵坐标，单位为标准化渲染像素。 */
  y: number;
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
};
