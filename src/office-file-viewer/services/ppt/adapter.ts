import {
  annotatePresentationImageReuse,
  markPresentationElementsPreviewable,
} from '../presentation/imagePreviewPolicy';
import type {
  PresentationDocument,
  SlideElement,
  SlideModel,
} from '../presentation/types';
import type { PptBinaryDocument, PptMasterModel, PptSlideModel } from './types';

function cloneMasterElement(
  element: SlideElement,
  slideId: string,
  index: number,
): SlideElement {
  return {
    ...element,
    id: `${slideId}-master-${index}-${element.id}`,
    zIndex: index,
  };
}

/** 将单张 PPT 幻灯片和对应母版合并为统一渲染模型。 */
export function adaptPptSlide(
  slide: PptSlideModel,
  masters: Map<number, PptMasterModel>,
): SlideModel {
  const masterElements = (
    masters.get(slide.masterId ?? Number.NaN)?.elements ?? []
  ).filter((element) => element.type !== 'text' || !element.placeholderType);
  // 母版标题/正文占位符只提供编辑提示和默认样式，不能作为放映正文继承到每一页。
  const inherited = markPresentationElementsPreviewable(
    masterElements.map((element, index) =>
      cloneMasterElement(element, slide.id, index),
    ),
    false,
  );
  const elements = annotatePresentationImageReuse([
    ...inherited,
    ...slide.elements.map((element, index) => ({
      ...element,
      zIndex: inherited.length + index,
    })),
  ]);
  return {
    id: slide.id,
    index: slide.index,
    width: slide.width,
    height: slide.height,
    hidden: slide.hidden,
    background: slide.background,
    speakerNotes: slide.speakerNotes,
    annotations: slide.annotations,
    transition: slide.transition,
    warnings: slide.warnings,
    elements,
  };
}

/** 将 PPT 私有结构适配为 PPTX 渲染器复用的统一演示文稿模型。 */
export function adaptPptDocument(
  source: PptBinaryDocument,
): PresentationDocument {
  return {
    width: source.width,
    height: source.height,
    theme: source.theme,
    slides: source.slides.map((slide) => adaptPptSlide(slide, source.masters)),
    warnings: source.warnings,
  };
}
