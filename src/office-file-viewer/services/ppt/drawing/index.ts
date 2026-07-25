// PPT 绘图解析入口，导出锚点、OfficeArt 属性和标准化形状能力。
export { parsePptDrawing } from './parsePptDrawing';
export { readPptOfficeArtProperties } from './readOfficeArtProperties';
export { readPptAnchor } from './readPptAnchor';
export type {
  PptOfficeArtProperty,
  PptShapeAnchor,
  PptShapeStyle,
} from './types';
