import { OFFICE_LARGE_FILE_THRESHOLDS } from '../performance/officePerformanceThresholds';
import type { SlideElement, SlideModel } from './types';

export const PRESENTATION_THUMBNAIL_VIRTUAL_SLIDE_COUNT = 50;
export const PRESENTATION_LAZY_SLIDE_COUNT = 80;
export const PRESENTATION_LAZY_ELEMENT_WEIGHT = 3000;
export const PRESENTATION_LAZY_SINGLE_SLIDE_ELEMENTS = 300;
export const PRESENTATION_PRELOAD_RADIUS = 2;

/** 描述 Presentation Viewer 的缩略图和幻灯片加载策略。 */
export type PresentationPerformanceProfile = {
  thumbnailMode: 'normal' | 'virtual';
  slideMode: 'materialized' | 'lazy';
  totalElementWeight: number;
};

/** 提前性能画像可使用的包目录和渐进解析统计。 */
export type PresentationPerformanceStats = {
  slideCount: number;
  totalElementWeight?: number;
  maxSlideElementCount?: number;
  mediaBytes?: number;
  compressedBytes?: number;
  uncompressedBytes?: number;
  mainXmlBytes?: number;
  singleMediaBytes?: number;
  cfbFileBytes?: number;
  cfbMainStreamBytes?: number;
  cfbResourceStreamBytes?: number;
  maxSlideParseMilliseconds?: number;
};

/** 按渲染成本估算单个元素权重，嵌套组会递归累计子元素。 */
export function getPresentationElementWeight(element: SlideElement): number {
  if (element.type === 'image') return 20;
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
