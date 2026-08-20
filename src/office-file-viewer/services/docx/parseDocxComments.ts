import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import { readXml } from '../../shared/ooxml/archive';
import {
  attr,
  descendantsByLocalName,
  parseXml,
  textContent,
} from '../../shared/ooxml/xml';
import type { OfficeAnnotation } from '../annotations/types';
import type {
  WordReviewDocument,
  WordReviewWarning,
  WordRevision,
  WordRevisionRecord,
} from '../word/review/types';

/** 批注范围端点在标准正文模型中的字符位置。 */
type DocxCommentPosition = {
  /** 正文块稳定标识。 */
  blockId: string;
  /** 当前块中的字符偏移。 */
  offset: number;
};

/** comments.xml 中不依赖正文范围的批注信息。 */
type DocxCommentDefinition = {
  /** OOXML 批注标识。 */
  id: string;
  /** 批注作者。 */
  author?: string;
  /** 批注创建日期。 */
  createdAt?: string;
  /** 批注正文纯文本。 */
  text: string;
  /** 批注正文首段标识，用于关联回复和解决状态。 */
  paragraphId?: string;
  /** 回复所属的父批注标识。 */
  parentId?: string;
  /** 批注线程是否已经解决。 */
  resolved?: boolean;
};

/** 正文解析期间持续补全的单条批注范围。 */
type DocxCommentRange = {
  /** commentRangeStart 对应的位置。 */
  start?: DocxCommentPosition;
  /** commentRangeEnd 对应的位置。 */
  end?: DocxCommentPosition;
  /** commentReference 对应的降级位置。 */
  reference?: DocxCommentPosition;
};

/** DOCX 正文和审阅部件共享的可变解析状态。 */
export type DocxReviewParseState = {
  /** 按源顺序排列的批注定义。 */
  comments: readonly DocxCommentDefinition[];
  /** 按批注标识索引的定义。 */
  commentById: ReadonlyMap<string, DocxCommentDefinition>;
  /** 正文扫描期间收集的批注范围。 */
  commentRanges: Map<string, DocxCommentRange>;
  /** 当前正文位置仍处于范围内的批注标识。 */
  activeCommentIds: Set<string>;
  /** 按稳定标识去重的修订记录。 */
  revisions: Map<string, WordRevision>;
  /** 解析降级诊断。 */
  warnings: WordReviewWarning[];
  /** 文档中第一个可用正文块，供损坏范围降级定位。 */
  firstBlockId?: string;
};

/** 按 XML 顺序提取批注段落文本，保留段落分隔。 */
function readCommentText(comment: Element) {
  const paragraphs = descendantsByLocalName(comment, 'p');
  if (!paragraphs.length) return textContent(comment).trim();
  return paragraphs
    .map((paragraph) => textContent(paragraph).trim())
    .filter(Boolean)
    .join('\n');
}

/** 读取 comments.xml 中的批注定义。 */
function readCommentDefinitions(entries: OfficeEntryMap) {
  const xml = readXml(entries, 'word/comments.xml');
  if (!xml) return [];
  const document = parseXml(xml);
  return descendantsByLocalName(document.documentElement, 'comment').map(
    (comment): DocxCommentDefinition => {
      const firstParagraph = descendantsByLocalName(comment, 'p')[0];
      return {
        id: attr(comment, 'w:id') ?? attr(comment, 'id') ?? '',
        author: attr(comment, 'w:author') ?? attr(comment, 'author'),
        createdAt: attr(comment, 'w:date') ?? attr(comment, 'date'),
        text: readCommentText(comment),
        paragraphId: firstParagraph
          ? attr(firstParagraph, 'w14:paraId') ?? attr(firstParagraph, 'paraId')
          : undefined,
      };
    },
  );
}

/** 读取 commentsExtended.xml 中的回复关系和解决状态。 */
function applyExtendedCommentMetadata(
  entries: OfficeEntryMap,
  comments: DocxCommentDefinition[],
) {
  const xml = readXml(entries, 'word/commentsExtended.xml');
  if (!xml) return;
  const paragraphIdToCommentId = new Map(
    comments.flatMap((comment) =>
      comment.paragraphId ? [[comment.paragraphId, comment.id] as const] : [],
    ),
  );
  const extendedByParagraphId = new Map(
    descendantsByLocalName(parseXml(xml).documentElement, 'commentEx').flatMap(
      (comment) => {
        const paragraphId =
          attr(comment, 'w15:paraId') ?? attr(comment, 'paraId');
        return paragraphId ? [[paragraphId, comment] as const] : [];
      },
    ),
  );

  comments.forEach((comment) => {
    if (!comment.paragraphId) return;
    const extended = extendedByParagraphId.get(comment.paragraphId);
    if (!extended) return;
    const parentParagraphId =
      attr(extended, 'w15:paraIdParent') ?? attr(extended, 'paraIdParent');
    comment.parentId = parentParagraphId
      ? paragraphIdToCommentId.get(parentParagraphId)
      : undefined;
    const done = attr(extended, 'w15:done') ?? attr(extended, 'done');
    comment.resolved = done === '1' || done === 'true';
  });
}

/** 建立 DOCX 审阅解析状态，正文解析只追加范围和修订。 */
export function createDocxReviewParseState(
  entries: OfficeEntryMap,
): DocxReviewParseState {
  const comments = readCommentDefinitions(entries).filter((comment) =>
    Boolean(comment.id),
  );
  applyExtendedCommentMetadata(entries, comments);
  return {
    comments,
    commentById: new Map(comments.map((comment) => [comment.id, comment])),
    commentRanges: new Map(),
    activeCommentIds: new Set(),
    revisions: new Map(),
    warnings: [],
  };
}

/** 记录当前文档首个正文块，供损坏批注范围安全降级。 */
export function registerDocxReviewBlock(
  state: DocxReviewParseState,
  blockId: string,
) {
  state.firstBlockId ??= blockId;
}

/** 开始追踪覆盖后续行内内容的批注范围。 */
export function startDocxCommentRange(
  state: DocxReviewParseState,
  id: string,
  position: DocxCommentPosition,
) {
  const range = state.commentRanges.get(id) ?? {};
  range.start = position;
  state.commentRanges.set(id, range);
  state.activeCommentIds.add(id);
}

/** 结束批注范围，并阻止后续内容继续继承该批注。 */
export function endDocxCommentRange(
  state: DocxReviewParseState,
  id: string,
  position: DocxCommentPosition,
) {
  const range = state.commentRanges.get(id) ?? {};
  range.end = position;
  state.commentRanges.set(id, range);
  state.activeCommentIds.delete(id);
}

/** 记录批注引用字符的位置，供缺失范围标记时降级定位。 */
export function recordDocxCommentReference(
  state: DocxReviewParseState,
  id: string,
  position: DocxCommentPosition,
) {
  const range = state.commentRanges.get(id) ?? {};
  range.reference = position;
  state.commentRanges.set(id, range);
}

/** 把当前范围内的批注标识复制到行内模型，避免泄露可变 Set。 */
export function getActiveDocxCommentIds(state: DocxReviewParseState) {
  return state.activeCommentIds.size
    ? [...state.activeCommentIds].filter((id) => state.commentById.has(id))
    : undefined;
}

/** 按稳定标识登记修订，重复引用不会重复增加摘要数量。 */
export function recordDocxRevision(
  state: DocxReviewParseState,
  revision: WordRevision,
) {
  state.revisions.set(revision.id, revision);
  return revision;
}

/** 为损坏或跨块批注范围选择不影响正文的降级目标。 */
function resolveAnnotationTarget(
  state: DocxReviewParseState,
  comment: DocxCommentDefinition,
) {
  const ownRange = state.commentRanges.get(comment.id);
  const parentRange = comment.parentId
    ? state.commentRanges.get(comment.parentId)
    : undefined;
  const range = ownRange ?? parentRange;
  const start = range?.start ?? range?.reference;
  const end = range?.end ?? range?.reference;
  const blockId = start?.blockId ?? end?.blockId ?? state.firstBlockId;
  if (!blockId) return undefined;

  if (!start || !end || start.blockId !== end.blockId) {
    state.warnings.push({
      code: 'DOCX_COMMENT_RANGE_DEGRADED',
      message: `DOCX 批注 ${comment.id} 的范围不完整，已定位到最近有效正文位置。`,
    });
  }
  const startOffset =
    start?.blockId === blockId ? start.offset : end?.offset ?? 0;
  const endOffset =
    end?.blockId === blockId
      ? Math.max(startOffset, end.offset)
      : Math.max(startOffset, startOffset + 1);
  return {
    kind: 'word-range' as const,
    blockId,
    startOffset,
    endOffset,
  };
}

/** 将正文范围与 comments.xml 定义合并为共享审阅模型。 */
export function buildDocxReviewDocument(
  state: DocxReviewParseState,
  noteCount = 0,
  revisionRecords: readonly WordRevisionRecord[] = [],
): WordReviewDocument | undefined {
  const annotations = state.comments.flatMap((comment): OfficeAnnotation[] => {
    const target = resolveAnnotationTarget(state, comment);
    return target
      ? [
          {
            id: comment.id,
            author: comment.author,
            createdAt: comment.createdAt,
            text: comment.text,
            resolved: comment.resolved,
            parentId: comment.parentId,
            target,
          },
        ]
      : [];
  });
  if (
    !annotations.length &&
    !state.revisions.size &&
    !noteCount &&
    !state.warnings.length
  ) {
    return undefined;
  }
  const revisionById = new Map(
    revisionRecords.map((revision) => [revision.id, revision] as const),
  );
  state.revisions.forEach((revision) => {
    if (revisionById.has(revision.id) || !state.firstBlockId) return;
    revisionById.set(revision.id, {
      ...revision,
      excerpt: '',
      target: {
        kind: 'word-range',
        blockId: state.firstBlockId,
        startOffset: 0,
        endOffset: 0,
      },
    });
  });
  const revisions = [...revisionById.values()];
  return {
    annotations,
    revisions,
    revisionCount: Math.max(state.revisions.size, revisions.length),
    noteCount: Math.max(0, Math.trunc(noteCount)),
    supportsRevisionModes: state.revisions.size > 0 || revisions.length > 0,
    warnings: [...state.warnings],
  };
}
