import type { SpreadsheetWarning } from '../../../spreadsheet/types';

/** 描述 XLS/BIFF8 解析使用的样式参数。 */
export type VectorStyle = {
  /** VectorStyle 的轮廓颜色；null 表示明确不绘制轮廓；未提供时沿用来源格式或渲染器的默认规则。 */
  stroke?: string;
  /** VectorStyle 的填充颜色、渐变或无填充标记；未提供时沿用来源格式或渲染器的默认规则。 */
  fill?: string;
  /** VectorStyle 的轮廓宽度，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  strokeWidth?: number;
  /** VectorStyle 的整体透明度，0 表示完全透明，1 表示完全不透明；未提供时沿用来源格式或渲染器的默认规则。 */
  opacity?: number;
  /** VectorStyle 的字体族名称；未提供时沿用来源格式或渲染器的默认规则。 */
  fontFamily?: string;
  /** VectorStyle 的字号，单位为标准化渲染像素；未提供时沿用来源格式或渲染器的默认规则。 */
  fontSize?: number;
  /** VectorStyle 的字体粗细值；未提供时沿用来源格式或渲染器的默认规则。 */
  fontWeight?: number;
  /** VectorStyle 的 textColor 文本值。 */
  textColor?: string;
};

/** 描述 VectorElement 在 XLS/BIFF8 解析中的数据结构。 */
export type VectorElement =
  | {
      /** 用于区分 VectorElement 不同结构分支的类型标识。 */
      type: 'line';
      /** VectorElement 在矢量图元坐标系中的 x1 几何值。 */
      x1: number;
      /** VectorElement 在矢量图元坐标系中的 y1 几何值。 */
      y1: number;
      /** VectorElement 在矢量图元坐标系中的 x2 几何值。 */
      x2: number;
      /** VectorElement 在矢量图元坐标系中的 y2 几何值。 */
      y2: number;
      /** VectorElement 使用的渲染或文本样式。 */
      style: VectorStyle;
    }
  | {
      /** 用于区分 VectorElement 不同结构分支的类型标识。 */
      type: 'polyline' | 'polygon';
      /** VectorElement 包含的 points 有序集合。 */
      points: Array<[number, number]>;
      /** VectorElement 使用的渲染或文本样式。 */
      style: VectorStyle;
    }
  | {
      /** 用于区分 VectorElement 不同结构分支的类型标识。 */
      type: 'rectangle' | 'ellipse';
      /** VectorElement 的 x 几何值，单位遵循对应 Office 二进制记录定义。 */
      x: number;
      /** VectorElement 的 y 几何值，单位遵循对应 Office 二进制记录定义。 */
      y: number;
      /** VectorElement 的 width 几何值，单位遵循对应 Office 二进制记录定义。 */
      width: number;
      /** VectorElement 的 height 几何值，单位遵循对应 Office 二进制记录定义。 */
      height: number;
      /** VectorElement 在矢量图元坐标系中的 radiusX 几何值。 */
      radiusX?: number;
      /** VectorElement 在矢量图元坐标系中的 radiusY 几何值。 */
      radiusY?: number;
      /** VectorElement 使用的渲染或文本样式。 */
      style: VectorStyle;
    }
  | {
      /** 用于区分 VectorElement 不同结构分支的类型标识。 */
      type: 'path';
      /** VectorElement 当前步骤需要处理的原始或标准化数据。 */
      data: string;
      /** VectorElement 使用的渲染或文本样式。 */
      style: VectorStyle;
    }
  | {
      /** 用于区分 VectorElement 不同结构分支的类型标识。 */
      type: 'text';
      /** VectorElement 的 x 几何值，单位遵循对应 Office 二进制记录定义。 */
      x: number;
      /** VectorElement 的 y 几何值，单位遵循对应 Office 二进制记录定义。 */
      y: number;
      /** VectorElement 携带或渲染的文本内容。 */
      text: string;
      /** VectorElement 使用的渲染或文本样式。 */
      style: VectorStyle;
    }
  | {
      /** 用于区分 VectorElement 不同结构分支的类型标识。 */
      type: 'image';
      /** VectorElement 的 x 几何值，单位遵循对应 Office 二进制记录定义。 */
      x: number;
      /** VectorElement 的 y 几何值，单位遵循对应 Office 二进制记录定义。 */
      y: number;
      /** VectorElement 的 width 几何值，单位遵循对应 Office 二进制记录定义。 */
      width: number;
      /** VectorElement 的 height 几何值，单位遵循对应 Office 二进制记录定义。 */
      height: number;
      /** VectorElement 的 dataUrl 文本值。 */
      dataUrl: string;
      /** VectorElement 使用的渲染或文本样式。 */
      style: VectorStyle;
    };

/** 描述 VectorScene 在 XLS/BIFF8 解析中的数据结构。 */
export type VectorScene = {
  /** VectorScene 的 width 几何值，单位遵循对应 Office 二进制记录定义。 */
  width: number;
  /** VectorScene 的 height 几何值，单位遵循对应 Office 二进制记录定义。 */
  height: number;
  /** VectorScene 关联的 viewBox 结构；字段形状由 [number, number, number, number] 定义。 */
  viewBox: [number, number, number, number];
  /** VectorScene 包含的 elements 有序集合。 */
  elements: VectorElement[];
  /** VectorScene 解析时产生但不阻止继续预览的警告集合。 */
  warnings: SpreadsheetWarning[];
};
