import type { SpreadsheetWarning } from '../../../spreadsheet/types';

/** WMF/EMF 矢量图元使用的绘制样式。 */
export type VectorStyle = {
  /** 轮廓颜色；null 表示明确不绘制轮廓。 */
  stroke?: string;
  /** 填充颜色、渐变或无填充标记。 */
  fill?: string;
  /** 轮廓宽度，单位为标准化渲染像素。 */
  strokeWidth?: number;
  /** 整体透明度，0 表示完全透明，1 表示完全不透明。 */
  opacity?: number;
  /** 字体族名称。 */
  fontFamily?: string;
  /** 字号，单位为标准化渲染像素。 */
  fontSize?: number;
  /** 字体粗细。 */
  fontWeight?: number;
  /** 文字使用的颜色。 */
  textColor?: string;
};

/** WMF/EMF 转换后支持的矢量图元联合类型。 */
export type VectorElement =
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'line';
      /** 在矢量图元坐标系中的 几何值。 */
      x1: number;
      /** 在矢量图元坐标系中的 几何值。 */
      y1: number;
      /** 在矢量图元坐标系中的 几何值。 */
      x2: number;
      /** 在矢量图元坐标系中的 几何值。 */
      y2: number;
      /** 当前内容使用的渲染样式。 */
      style: VectorStyle;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'polyline' | 'polygon';
      /** 折线或多边形的顶点坐标。 */
      points: Array<[number, number]>;
      /** 当前内容使用的渲染样式。 */
      style: VectorStyle;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'rectangle' | 'ellipse';
      /** 相对定位区域左侧的横坐标，单位为标准化渲染像素。 */
      x: number;
      /** 相对定位区域顶部的纵坐标，单位为标准化渲染像素。 */
      y: number;
      /** 宽度，单位为标准化渲染像素。 */
      width: number;
      /** 高度，单位为标准化渲染像素。 */
      height: number;
      /** 圆角或椭圆在水平方向的半径。 */
      radiusX?: number;
      /** 圆角或椭圆在垂直方向的半径。 */
      radiusY?: number;
      /** 当前内容使用的渲染样式。 */
      style: VectorStyle;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'path';
      /** SVG 路径命令文本。 */
      data: string;
      /** 当前内容使用的渲染样式。 */
      style: VectorStyle;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'text';
      /** 相对定位区域左侧的横坐标，单位为标准化渲染像素。 */
      x: number;
      /** 相对定位区域顶部的纵坐标，单位为标准化渲染像素。 */
      y: number;
      /** 文本内容。 */
      text: string;
      /** 当前内容使用的渲染样式。 */
      style: VectorStyle;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'image';
      /** 相对定位区域左侧的横坐标，单位为标准化渲染像素。 */
      x: number;
      /** 相对定位区域顶部的纵坐标，单位为标准化渲染像素。 */
      y: number;
      /** 宽度，单位为标准化渲染像素。 */
      width: number;
      /** 高度，单位为标准化渲染像素。 */
      height: number;
      /** 内嵌位图使用的数据地址。 */
      dataUrl: string;
      /** 当前内容使用的渲染样式。 */
      style: VectorStyle;
    };

/** WMF/EMF 转换后的画布、图元和警告。 */
export type VectorScene = {
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 矢量场景的最小横纵坐标及宽高。 */
  viewBox: [number, number, number, number];
  /** 按绘制顺序排列的矢量图元。 */
  elements: VectorElement[];
  /** 解析时产生但不阻止继续预览的警告。 */
  warnings: SpreadsheetWarning[];
};
