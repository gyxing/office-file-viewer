import type { OfficeAnnotation } from '../../services/annotations/types';
import type {
  WordRevisionKind,
  WordRevisionRecord,
} from '../../services/word/review/types';
import type { WordMarkupCalloutAnchor } from './wordMarkupCalloutLayout';

/** 浏览器视口与 Word 滚动容器相交后的可见边界。 */
export type VisibleScrollerBoundary = Readonly<{
  /** 可见区域左边界。 */
  left: number;
  /** 可见区域上边界。 */
  top: number;
  /** 可见区域右边界。 */
  right: number;
  /** 可见区域下边界。 */
  bottom: number;
}>;

/** 单行批注标记的高度。 */
const WORD_COMMENT_CALLOUT_HEIGHT = 56;
/** 多行批注标记的高度。 */
const WORD_COMMENT_CALLOUT_LONG_HEIGHT = 72;
/** 单行修订标记的高度。 */
const WORD_REVISION_CALLOUT_HEIGHT = 48;
/** 多行修订标记的高度。 */
const WORD_REVISION_CALLOUT_LONG_HEIGHT = 60;

/** 校验 DOM 中恢复出的修订类别。 */
function isWordRevisionKind(
  value: string | undefined,
): value is WordRevisionKind {
  return (
    value === 'insert' ||
    value === 'delete' ||
    value === 'move-from' ||
    value === 'move-to' ||
    value === 'format'
  );
}

/** 判断正文标记是否进入当前滚动视口。 */
export function intersectsBoundary(
  rect: DOMRect,
  boundary: VisibleScrollerBoundary,
) {
  return (
    rect.bottom >= boundary.top &&
    rect.top <= boundary.bottom &&
    rect.right >= boundary.left &&
    rect.left <= boundary.right
  );
}

/** 选择元素最后一个可见行盒，让连接线对准文字范围末端。 */
function getLastVisibleRect(
  element: HTMLElement,
  boundary: VisibleScrollerBoundary,
) {
  const visibleRects = Array.from(element.getClientRects()).filter((rect) =>
    intersectsBoundary(rect, boundary),
  );
  return visibleRects[visibleRects.length - 1];
}

/** 收集当前页面可见批注标识。 */
export function collectVisibleAnnotationIds(
  article: HTMLElement,
  boundary: VisibleScrollerBoundary,
) {
  return [
    ...new Set(
      Array.from(
        article.querySelectorAll<HTMLElement>('[data-office-annotation-id]'),
      ).flatMap((element) => {
        const id = element.dataset.officeAnnotationId;
        return id && getLastVisibleRect(element, boundary) ? [id] : [];
      }),
    ),
  ];
}

/** 把当前视口内的批注范围转换为页侧标记锚点。 */
export function collectVisibleCommentAnchors({
  article,
  boundary,
  contentOriginLeft,
  contentOriginTop,
  annotationById,
}: {
  /** 当前 Word 页面。 */
  article: HTMLElement;
  /** 当前滚动容器可见边界。 */
  boundary: VisibleScrollerBoundary;
  /** 滚动内容坐标系在浏览器视口中的横向原点。 */
  contentOriginLeft: number;
  /** 滚动内容坐标系在浏览器视口中的纵向原点。 */
  contentOriginTop: number;
  /** 已按稳定标识加载的批注元数据。 */
  annotationById: ReadonlyMap<string, OfficeAnnotation>;
}): WordMarkupCalloutAnchor[] {
  const anchorRectById = new Map<string, DOMRect>();
  article
    .querySelectorAll<HTMLElement>('[data-office-annotation-id]')
    .forEach((element) => {
      const id = element.dataset.officeAnnotationId;
      const rect = getLastVisibleRect(element, boundary);
      if (id && rect) anchorRectById.set(id, rect);
    });
  return Array.from(anchorRectById.entries()).flatMap(([id, rect]) => {
    const annotation = annotationById.get(id);
    if (!annotation) return [];
    const excerpt = annotation.text.trim();
    return [
      {
        key: `comment:${id}`,
        id,
        type: 'comment' as const,
        author: annotation.author,
        createdAt: annotation.createdAt,
        excerpt,
        resolved: annotation.resolved,
        left: rect.left - contentOriginLeft,
        top: rect.top - contentOriginTop,
        right: rect.right - contentOriginLeft,
        bottom: rect.bottom - contentOriginTop,
        height:
          excerpt.length > 32
            ? WORD_COMMENT_CALLOUT_LONG_HEIGHT
            : WORD_COMMENT_CALLOUT_HEIGHT,
      },
    ];
  });
}

/** 判断两个修订锚点是否覆盖同一段正文区域。 */
function overlapsRevisionAnchor(
  first: WordMarkupCalloutAnchor,
  second: WordMarkupCalloutAnchor,
) {
  return (
    first.left < second.right &&
    first.right > second.left &&
    first.top < second.bottom &&
    first.bottom > second.top
  );
}

/** 合并 Word 为同一次修改生成的段落级和行内级重复修订。 */
function deduplicateNestedRevisions(
  anchors: readonly WordMarkupCalloutAnchor[],
) {
  const deduplicated: WordMarkupCalloutAnchor[] = [];
  anchors.forEach((anchor) => {
    const duplicateIndex = deduplicated.findIndex(
      (existing) =>
        existing.type === 'revision' &&
        existing.revisionKind === anchor.revisionKind &&
        existing.author === anchor.author &&
        existing.createdAt === anchor.createdAt &&
        overlapsRevisionAnchor(existing, anchor),
    );
    if (duplicateIndex < 0) {
      deduplicated.push(anchor);
      return;
    }
    const existing = deduplicated[duplicateIndex];
    const existingArea =
      (existing.right - existing.left) * (existing.bottom - existing.top);
    const anchorArea =
      (anchor.right - anchor.left) * (anchor.bottom - anchor.top);
    if (anchorArea < existingArea) deduplicated[duplicateIndex] = anchor;
  });
  return deduplicated;
}

/** 合并同一修订的可见片段，并转换为页侧标记锚点。 */
export function collectVisibleRevisionAnchors({
  article,
  boundary,
  contentOriginLeft,
  contentOriginTop,
  activeRevision,
}: {
  /** 当前 Word 页面。 */
  article: HTMLElement;
  /** 当前滚动容器可见边界。 */
  boundary: VisibleScrollerBoundary;
  /** 滚动内容坐标系在浏览器视口中的横向原点。 */
  contentOriginLeft: number;
  /** 滚动内容坐标系在浏览器视口中的纵向原点。 */
  contentOriginTop: number;
  /** 当前由正文点击激活的修订。 */
  activeRevision?: WordRevisionRecord;
}): WordMarkupCalloutAnchor[] {
  const groupedAnchors = new Map<
    string,
    { element: HTMLElement; rect: DOMRect; excerpts: Set<string> }
  >();
  article
    .querySelectorAll<HTMLElement>('[data-office-word-revision]')
    .forEach((element) => {
      const id = element.dataset.officeWordRevision;
      const kind = element.dataset.officeWordRevisionKind;
      const rect = getLastVisibleRect(element, boundary);
      if (!id || !isWordRevisionKind(kind) || !rect) return;
      const excerpt = (element.textContent ?? '').replace(/\s+/g, ' ').trim();
      const current = groupedAnchors.get(id);
      if (current) {
        current.element = element;
        current.rect = rect;
        if (excerpt) current.excerpts.add(excerpt);
        return;
      }
      groupedAnchors.set(id, {
        element,
        rect,
        excerpts: new Set(excerpt ? [excerpt] : []),
      });
    });
  const anchors = Array.from(groupedAnchors.entries()).flatMap(
    ([id, { element, rect, excerpts }]) => {
      const revisionKind = element.dataset.officeWordRevisionKind;
      if (!isWordRevisionKind(revisionKind)) return [];
      const record = activeRevision?.id === id ? activeRevision : undefined;
      const excerpt = record?.excerpt || Array.from(excerpts).join(' ');
      return [
        {
          key: `revision:${id}`,
          id,
          type: 'revision' as const,
          revisionKind,
          author: record?.author ?? element.dataset.officeWordRevisionAuthor,
          createdAt:
            record?.createdAt ?? element.dataset.officeWordRevisionCreatedAt,
          excerpt,
          left: rect.left - contentOriginLeft,
          top: rect.top - contentOriginTop,
          right: rect.right - contentOriginLeft,
          bottom: rect.bottom - contentOriginTop,
          height:
            excerpt.length > 28
              ? WORD_REVISION_CALLOUT_LONG_HEIGHT
              : WORD_REVISION_CALLOUT_HEIGHT,
        },
      ];
    },
  );
  return deduplicateNestedRevisions(anchors);
}
