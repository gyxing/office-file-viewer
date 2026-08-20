import { resolvePackageMediaRef } from '../../shared/ooxml/media';
import { emuToPx } from '../../shared/ooxml/units';
import {
  attr,
  childByLocalName,
  descendantByLocalName,
} from '../../shared/ooxml/xml';
import {
  getPresentationMediaMimeType,
  type PresentationMediaKind,
} from '../presentation/mediaTypes';
import type { MediaElement } from '../presentation/types';
import type { PptxPackageState } from './PptxPackageContext';

function emuValue(node: Element | null, name: string) {
  const value = attr(node, name);
  return value === undefined ? undefined : emuToPx(Number(value));
}

function resolveMediaRef(target: string, packageState: PptxPackageState) {
  return resolvePackageMediaRef(
    target,
    packageState.mediaByPath,
    packageState.mediaByName,
    'ppt',
  );
}

function detectMediaKind(node: Element, target: string): PresentationMediaKind {
  if (descendantByLocalName(node, 'videoFile')) return 'video';
  if (descendantByLocalName(node, 'audioFile')) return 'audio';
  const mime = getPresentationMediaMimeType(target);
  return mime.startsWith('audio/') ? 'audio' : 'video';
}

/** 从图片形状的 media、audioFile 或 videoFile 关系恢复可播放媒体。 */
export function parsePptxMediaElement(
  node: Element,
  index: number,
  packageState: PptxPackageState,
  relationships: PptxPackageState['relationships'][string],
): MediaElement | undefined {
  const mediaNode =
    descendantByLocalName(node, 'media') ??
    descendantByLocalName(node, 'videoFile') ??
    descendantByLocalName(node, 'audioFile') ??
    descendantByLocalName(node, 'wavAudioFile');
  if (!mediaNode) return undefined;
  const relationshipId =
    attr(mediaNode, 'r:embed') ??
    attr(mediaNode, 'embed') ??
    attr(mediaNode, 'r:link') ??
    attr(mediaNode, 'link');
  const relationship = relationshipId
    ? relationships[relationshipId]
    : undefined;
  if (!relationship?.target) return undefined;
  const sourceKind =
    relationship.targetMode?.toLowerCase() === 'external'
      ? 'external'
      : 'embedded';
  const source =
    sourceKind === 'external'
      ? relationship.target
      : resolveMediaRef(relationship.target, packageState);
  if (!source) return undefined;

  const xfrm = childByLocalName(childByLocalName(node, 'spPr') ?? node, 'xfrm');
  const blip = descendantByLocalName(node, 'blip');
  const posterRelationshipId = attr(blip, 'r:embed') ?? attr(blip, 'embed');
  const posterTarget = posterRelationshipId
    ? relationships[posterRelationshipId]?.target
    : undefined;
  const properties = descendantByLocalName(node, 'cNvPr');
  const sourceObjectId = attr(properties, 'id') ?? undefined;
  return {
    id: `media-${index}`,
    sourceObjectId,
    type: 'media',
    x: emuValue(childByLocalName(xfrm, 'off'), 'x') ?? 0,
    y: emuValue(childByLocalName(xfrm, 'off'), 'y') ?? 0,
    width: emuValue(childByLocalName(xfrm, 'ext'), 'cx') ?? 0,
    height: emuValue(childByLocalName(xfrm, 'ext'), 'cy') ?? 0,
    rotate: attr(xfrm, 'rot') ? Number(attr(xfrm, 'rot')) / 60000 : undefined,
    flipH: attr(xfrm, 'flipH') === '1',
    flipV: attr(xfrm, 'flipV') === '1',
    posterSrc: posterTarget
      ? resolveMediaRef(posterTarget, packageState)
      : undefined,
    alt:
      attr(properties, 'descr') ??
      attr(properties, 'name') ??
      relationship.target.split('/').pop(),
    media: {
      kind: detectMediaKind(node, relationship.target),
      sourceKind,
      source,
      mimeType: getPresentationMediaMimeType(relationship.target),
      fileName: relationship.target.split('/').pop(),
      loop: attr(mediaNode, 'loop') === '1',
    },
  };
}
