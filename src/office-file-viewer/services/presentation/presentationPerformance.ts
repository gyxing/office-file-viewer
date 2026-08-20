import { OFFICE_LARGE_FILE_THRESHOLDS } from '../performance/officePerformanceThresholds';
import type { SlideElement, SlideModel } from './types';

/** 启用缩略图虚拟列表的幻灯片数量阈值。 */
export const PRESENTATION_THUMBNAIL_VIRTUAL_SLIDE_COUNT = 50;
/** 启用幻灯片按需加载的页数阈值。 */
export const PRESENTATION_LAZY_SLIDE_COUNT = 80;
/** 启用幻灯片按需加载的累计元素权重阈值。 */
export const PRESENTATION_LAZY_ELEMENT_WEIGHT = 3000;
/** 单页元素过多时启用按需加载的数量阈值。 */
export const PRESENTATION_LAZY_SINGLE_SLIDE_ELEMENTS = 300;
/** 当前幻灯片前后预加载的页数。 */
export const PRESENTATION_PRELOAD_RADIUS = 2;

/** 演示文稿 性能 性能档案。 */
export type PresentationPerformanceProfile = {
  /** 缩略图列表采用普通渲染还是虚拟渲染。 */
  thumbnailMode: 'normal' | 'virtual';
  /** 幻灯片采用全部物化还是按需加载。 */
  slideMode: 'materialized' | 'lazy';
  /** 全部内容元素折算后的累计渲染权重。 */
  totalElementWeight: number;
};

/** 提前性能画像可使用的包目录和渐进解析统计。 */
export type PresentationPerformanceStats = {
  /** 演示文稿包含的幻灯片数量。 */
  slideCount: number;
  /** 全部内容元素折算后的累计渲染权重。 */
  totalElementWeight?: number;
  /** 单张幻灯片包含的最大元素数量。 */
  maxSlideElementCount?: number;
  /** 媒体资源累计大小，单位为字节。 */
  mediaBytes?: number;
  /** 源压缩包大小，单位为字节。 */
  compressedBytes?: number;
  /** 源压缩包解压后的累计大小，单位为字节。 */
  uncompressedBytes?: number;
  /** 主 XML 文档大小，单位为字节。 */
  mainXmlBytes?: number;
  /** 单个最大媒体资源的大小，单位为字节。 */
  singleMediaBytes?: number;
  /** CFB 文件总大小，单位为字节。 */
  cfbFileBytes?: number;
  /** CFB 主数据流大小，单位为字节。 */
  cfbMainStreamBytes?: number;
  /** CFB 最大资源流大小，单位为字节。 */
  cfbResourceStreamBytes?: number;
  /** 单张幻灯片解析耗时的最大值，单位为毫秒。 */
  maxSlideParseMilliseconds?: number;
};

/** 按渲染成本估算单个元素权重，嵌套组会递归累计子元素。 */
export function getPresentationElementWeight(element: SlideElement): number {
  if (element.type === 'image') return 20;
  if (element.type === 'media') return 30;
  if (element.type === 'chart') return 30;
  if (element.type === 'table') return Math.max(1, element.rows.length) * 4;
  if (element.type === 'group') {
    return element.children.reduce(
      (total, child) => total + getPresentationElementWeight(child),
      0,
    );
  }
  return 1;
}

/** 汇总已物化幻灯片的元素权重。 */
export function getPresentationSlideWeight(slide: SlideModel) {
  return slide.elements.reduce(
    (total, element) => total + getPresentationElementWeight(element),
    0,
  );
}

/** 根据共享文件阈值和 Presentation 专属阈值生成单向性能模式。 */
export function createPresentationPerformanceProfile(
  stats: PresentationPerformanceStats,
): PresentationPerformanceProfile {
  const totalElementWeight = stats.totalElementWeight ?? 0;
  const useLazySlides =
    stats.slideCount >= PRESENTATION_LAZY_SLIDE_COUNT ||
    totalElementWeight >= PRESENTATION_LAZY_ELEMENT_WEIGHT ||
    (stats.maxSlideElementCount ?? 0) >=
      PRESENTATION_LAZY_SINGLE_SLIDE_ELEMENTS ||
    (stats.mediaBytes ?? 0) >=
      OFFICE_LARGE_FILE_THRESHOLDS.ooxmlUncompressedBytes ||
    (stats.compressedBytes ?? 0) >=
      OFFICE_LARGE_FILE_THRESHOLDS.ooxmlCompressedBytes ||
    (stats.uncompressedBytes ?? 0) >=
      OFFICE_LARGE_FILE_THRESHOLDS.ooxmlUncompressedBytes ||
    (stats.mainXmlBytes ?? 0) >=
      OFFICE_LARGE_FILE_THRESHOLDS.ooxmlMainXmlBytes ||
    (stats.singleMediaBytes ?? 0) >=
      OFFICE_LARGE_FILE_THRESHOLDS.ooxmlSingleMediaBytes ||
    (stats.cfbFileBytes ?? 0) >= OFFICE_LARGE_FILE_THRESHOLDS.cfbFileBytes ||
    (stats.cfbMainStreamBytes ?? 0) >=
      OFFICE_LARGE_FILE_THRESHOLDS.cfbMainStreamBytes ||
    (stats.cfbResourceStreamBytes ?? 0) >=
      OFFICE_LARGE_FILE_THRESHOLDS.cfbResourceStreamBytes ||
    (stats.maxSlideParseMilliseconds ?? 0) >=
      OFFICE_LARGE_FILE_THRESHOLDS.slowTaskMilliseconds;

  return {
    thumbnailMode:
      stats.slideCount > PRESENTATION_THUMBNAIL_VIRTUAL_SLIDE_COUNT
        ? 'virtual'
        : 'normal',
    slideMode: useLazySlides ? 'lazy' : 'materialized',
    totalElementWeight,
  };
}

/** 从完整文稿模型生成运行期性能画像。 */
export function profileMaterializedPresentation(slides: readonly SlideModel[]) {
  const weights = slides.map(getPresentationSlideWeight);
  return createPresentationPerformanceProfile({
    slideCount: slides.length,
    totalElementWeight: weights.reduce((total, weight) => total + weight, 0),
    maxSlideElementCount: Math.max(
      0,
      ...slides.map((slide) => slide.elements.length),
    ),
  });
}
