import { emuToPx } from '../../shared/ooxml/units';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  descendantsByLocalName,
  parseXml,
} from '../../shared/ooxml/xml';
import type {
  PresentationAnnotation,
  PresentationWarning,
  SlideElement,
} from '../presentation/types';
import type { PptxPackageState } from './PptxPackageContext';

/** 单页 PPTX 批注解析结果。 */
export type PptxCommentParseResult = Readonly<{
  /** 当前页可展示和导航的批注。 */
  annotations: PresentationAnnotation[];
  /** 未知或损坏批注部件的局部降级说明。 */
  warnings: PresentationWarning[];
}>;

type PptxCommentAuthor = Readonly<{
  name?: string;
  initials?: string;
}>;

const authorCache = new WeakMap<
  PptxPackageState,
  ReadonlyMap<string, PptxCommentAuthor>
>();

function readCommentAuthors(packageState: PptxPackageState) {
  const cached = authorCache.get(packageState);
  if (cached) return cached;
  const authors = new Map<string, PptxCommentAuthor>();
  packageState.entries.forEach((value, path) => {
    if (
      typeof value !== 'string' ||
      !/commentauthors|\/authors\//i.test(path)
    ) {
      return;
    }
    try {
      const root = parseXml(value).documentElement;
      ['cmAuthor', 'author'].forEach((name) => {
        descendantsByLocalName(root, name).forEach((node) => {
          const id = attr(node, 'id');
          if (!id) return;
          authors.set(id, {
            name: attr(node, 'name') ?? undefined,
            initials: attr(node, 'initials') ?? undefined,
          });
        });
      });
    } catch {
      // 单个作者扩展损坏时仍允许使用批注正文中的匿名作者。
    }
  });
  authorCache.set(packageState, authors);
  return authors;
}

function readCommentText(node: Element) {
  const plainText = childByLocalName(node, 'text')?.textContent?.trim();
  if (plainText) return plainText;
  const textBody = childByLocalName(node, 'txBody');
  return descendantsByLocalName(textBody, 't')
    .map((item) => item.textContent ?? '')
    .join('')
    .trim();
}

function normalizeCommentCoordinate(value: string | undefined, size: number) {
  if (value === undefined) return undefined;
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return undefined;
  return Math.abs(parsed) > Math.max(10000, size * 8)
    ? emuToPx(parsed)
    : parsed;
}

function findSourceObjectId(node: Element) {
  const anchors = Array.from(node.getElementsByTagName('*')).filter((item) =>
    /mk$/i.test(item.localName),
  );
  for (const anchor of anchors) {
    const value =
      attr(anchor, 'spid') ?? attr(anchor, 'shapeId') ?? attr(anchor, 'objId');
    if (value) return value;
  }
  return undefined;
}

function findTargetElement(
  node: Element,
  elements: readonly SlideElement[],
  x?: number,
  y?: number,
) {
  const sourceObjectId = findSourceObjectId(node);
  const explicit = sourceObjectId
    ? elements.find((element) => element.sourceObjectId === sourceObjectId)
    : undefined;
  if (explicit) return explicit.id;
  if (x === undefined || y === undefined || !elements.length) return undefined;
  return elements
    .map((element) => {
      const within =
        x >= element.x &&
        x <= element.x + element.width &&
        y >= element.y &&
        y <= element.y + element.height;
      const dx = x - (element.x + element.width / 2);
      const dy = y - (element.y + element.height / 2);
      return { id: element.id, distance: within ? -1 : Math.hypot(dx, dy) };
    })
    .sort((left, right) => left.distance - right.distance)[0]?.id;
}

function createAnnotation(
  node: Element,
  slideId: string,
  slideIndex: number,
  width: number,
  height: number,
  elements: readonly SlideElement[],
  authors: ReadonlyMap<string, PptxCommentAuthor>,
  ordinal: number,
  parentId?: string,
): PresentationAnnotation | undefined {
  const text = readCommentText(node);
  if (!text) return undefined;
  const authorId = attr(node, 'authorId') ?? '';
  const commentId = attr(node, 'id') ?? attr(node, 'idx') ?? String(ordinal);
  const id = `pptx-comment-${slideIndex + 1}-${commentId}`;
  const pos = childByLocalName(node, 'pos');
  const x = normalizeCommentCoordinate(attr(pos, 'x'), width);
  const y = normalizeCommentCoordinate(attr(pos, 'y'), height);
  return {
    id,
    author: authors.get(authorId)?.name,
    createdAt: attr(node, 'created') ?? attr(node, 'dt') ?? undefined,
    text,
    resolved: attr(node, 'status')?.toLowerCase() === 'resolved',
    parentId,
    x,
    y,
    target: {
      kind: 'presentation-element',
      slideId,
      slideIndex,
      elementId: findTargetElement(node, elements, x, y),
    },
  };
}

/** 解析传统与现代 PPTX 批注关系，未知扩展只产生局部 warning。 */
export function parsePptxComments(
  packageState: PptxPackageState,
  relsPath: string,
  slideId: string,
  slideIndex: number,
  width: number,
  height: number,
  elements: readonly SlideElement[] = [],
): PptxCommentParseResult {
  const annotations: PresentationAnnotation[] = [];
  const warnings: PresentationWarning[] = [];
  const authors = readCommentAuthors(packageState);
  const targets = Object.values(packageState.relationships[relsPath] ?? {})
    .filter(
      (relationship) =>
        relationship.type?.toLowerCase().endsWith('/comments') ||
        /\/comments\//i.test(relationship.target),
    )
    .map((relationship) => relationship.target);

  targets.forEach((target) => {
    const xml = packageState.entries.get(target);
    if (typeof xml !== 'string') return;
    try {
      const root = parseXml(xml).documentElement;
      const comments = childrenByLocalName(root, 'cm');
      comments.forEach((node, index) => {
        const annotation = createAnnotation(
          node,
          slideId,
          slideIndex,
          width,
          height,
          elements,
          authors,
          index,
        );
        if (!annotation) return;
        annotations.push(annotation);
        const replies = childrenByLocalName(
          childByLocalName(node, 'replyLst'),
          'reply',
        );
        replies.forEach((reply, replyIndex) => {
          const parsed = createAnnotation(
            reply,
            slideId,
            slideIndex,
            width,
            height,
            elements,
            authors,
            comments.length + index * 1000 + replyIndex,
            annotation.id,
          );
          if (parsed) annotations.push(parsed);
        });
      });
    } catch {
      warnings.push({
        code: 'PPTX_COMMENT_PART_CORRUPT',
        message: '当前幻灯片的一组批注无法读取，主体内容已继续显示',
        slideIndex,
      });
    }
  });
  return { annotations, warnings };
}
