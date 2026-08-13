import { createOfficeHyperlinkFromTarget } from '../../shared/hyperlink/createOfficeHyperlink';
import type {
  OfficeHyperlink,
  OfficePresentationHyperlinkTarget,
} from '../../shared/hyperlink/types';
import type { OfficeRelationship } from '../../shared/ooxml/media';
import { attr, childByLocalName } from '../../shared/ooxml/xml';

/** PPTX 关系目标路径到零基幻灯片索引的映射。 */
export type PptxSlideTargetMap = Readonly<Record<string, number>>;

function presentationHyperlink(
  target: OfficePresentationHyperlinkTarget,
  screenTip?: string,
): OfficeHyperlink {
  return {
    kind: 'internal',
    target: { family: 'presentation', ...target },
    screenTip,
  };
}

function readShowJump(action: string) {
  const jump = /[?&]jump=([^&]+)/i.exec(action)?.[1]?.toLowerCase();
  return jump === 'firstslide'
    ? 'first'
    : jump === 'lastslide'
    ? 'last'
    : jump === 'nextslide'
    ? 'next'
    : jump === 'previousslide'
    ? 'previous'
    : undefined;
}

function resolveSlideIndex(
  relationship: OfficeRelationship | undefined,
  slideTargets?: PptxSlideTargetMap,
) {
  if (!relationship?.target) return undefined;
  const mapped = slideTargets?.[relationship.target];
  if (mapped !== undefined) return mapped;
  const fileIndex = /\/slide(\d+)\.xml$/i.exec(relationship.target)?.[1];
  return fileIndex ? Number(fileIndex) - 1 : undefined;
}

/** 判断非可视属性中的链接是否带 PowerPoint 放映动作。 */
export function hasPptxHyperlinkAction(node: Element | null) {
  return Boolean(attr(childByLocalName(node, 'hlinkClick'), 'action'));
}

/** 解析 run 属性或对象非可视属性中的 a:hlinkClick。 */
export function parsePptxHyperlink(
  container: Element | null,
  relationships: Record<string, OfficeRelationship>,
  slideTargets?: PptxSlideTargetMap,
): OfficeHyperlink | undefined {
  const click = childByLocalName(container, 'hlinkClick');
  if (!click) return undefined;
  const screenTip = attr(click, 'tooltip');
  const action = attr(click, 'action')?.toLowerCase() ?? '';
  const relationshipId = attr(click, 'r:id') ?? attr(click, 'id');
  const relationship = relationshipId
    ? relationships[relationshipId]
    : undefined;

  if (action.includes('hlinkshowjump')) {
    const jump = readShowJump(action);
    return jump
      ? presentationHyperlink({ action: jump }, screenTip)
      : undefined;
  }
  if (action.includes('hlinksldjump')) {
    const slideIndex = resolveSlideIndex(relationship, slideTargets);
    return slideIndex === undefined
      ? undefined
      : presentationHyperlink({ slideIndex }, screenTip);
  }
  if (/macro|program|ole|sound/i.test(action)) return undefined;
  if (relationship?.type?.toLowerCase().endsWith('/slide')) {
    const slideIndex = resolveSlideIndex(relationship, slideTargets);
    return slideIndex === undefined
      ? undefined
      : presentationHyperlink({ slideIndex }, screenTip);
  }
  return relationship?.target
    ? createOfficeHyperlinkFromTarget(relationship.target, screenTip)
    : undefined;
}
