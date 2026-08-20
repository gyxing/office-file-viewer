import type { OfficeAnnotation } from '../annotations/types';
import { WordRevisionRecordCollector } from '../word/review/WordRevisionRecordCollector';
import type { WordReviewDocument } from '../word/review/types';
import type { DocBinaryContent, DocFib, DocTextSegment } from './docParseTypes';
import { readDocStorySegments } from './readDocBinaryContent';
import type { DocBlock, DocTextInline } from './types';

/** PlcfandRef 与批注 story 合并后的轻量定义。 */
export type DocCommentDefinition = {
  /** 当前文档中的稳定批注标识。 */
  id: string;
  /** 作者名称或 ATRDPre10 中的作者缩写。 */
  author?: string;
  /** 批注正文纯文本。 */
  text: string;
  /** 批注引用字符在主文档中的 CP。 */
  charPosition: number;
};

/** DOC/WPS 批注结构解析结果。 */
export type DocCommentParseResult = {
  /** 按源文档顺序排列的批注定义。 */
  comments: DocCommentDefinition[];
  /** 无法恢复完整范围时生成的稳定警告。 */
  warnings: string[];
};

/** 读取批注作者名称数组中的连续 XST 字符串。 */
function readCommentOwners(tableStream: Uint8Array, fib: DocFib) {
  const start = fib.fcGrpXstAtnOwners;
  const end = start + fib.lcbGrpXstAtnOwners;
  if (!start || end > tableStream.length) return [];
  const view = new DataView(
    tableStream.buffer,
    tableStream.byteOffset,
    tableStream.byteLength,
  );
  const owners: string[] = [];
  let offset = start;
  while (offset + 2 <= end) {
    const length = view.getUint16(offset, true);
    offset += 2;
    const byteLength = length * 2;
    if (length > 55 || offset + byteLength > end) break;
    owners.push(
      new TextDecoder('utf-16le')
        .decode(tableStream.slice(offset, offset + byteLength))
        .replace(/\u0000+$/g, ''),
    );
    offset += byteLength;
  }
  return owners;
}

/** 从 ATRDPre10 的固定 20 字节缓冲读取作者缩写。 */
function readCommentInitials(record: Uint8Array) {
  if (record.length < 20) return undefined;
  const view = new DataView(
    record.buffer,
    record.byteOffset,
    record.byteLength,
  );
  const length = Math.min(9, view.getUint16(0, true));
  if (!length) return undefined;
  return new TextDecoder('utf-16le')
    .decode(record.slice(2, 2 + length * 2))
    .replace(/\u0000+$/g, '');
}

/** 读取 PlcfandRef 的 CP、作者索引和作者缩写。 */
function readCommentReferences(tableStream: Uint8Array, fib: DocFib) {
  const offset = fib.fcPlcfandRef;
  const length = fib.lcbPlcfandRef;
  if (!offset || length < 38 || offset + length > tableStream.length) return [];
  const count = Math.floor((length - 4) / 34);
  if (!count) return [];
  const view = new DataView(
    tableStream.buffer,
    tableStream.byteOffset,
    tableStream.byteLength,
  );
  const recordsOffset = offset + (count + 1) * 4;
  return Array.from({ length: count }, (_, index) => {
    const recordOffset = recordsOffset + index * 30;
    const record = tableStream.slice(recordOffset, recordOffset + 30);
    return {
      charPosition: view.getUint32(offset + index * 4, true),
      authorIndex:
        record.length >= 22
          ? new DataView(
              record.buffer,
              record.byteOffset,
              record.byteLength,
            ).getUint16(20, true)
          : -1,
      initials: readCommentInitials(record),
    };
  });
}

/** 读取 PlcfandTxt 中与批注数量对应的正文边界。 */
function readCommentTextBoundaries(
  tableStream: Uint8Array,
  fib: DocFib,
  commentCount: number,
) {
  const offset = fib.fcPlcfandTxt;
  const length = fib.lcbPlcfandTxt;
  if (!offset || length < 8 || offset + length > tableStream.length) return [];
  const availableCount = Math.floor(length / 4);
  const count = Math.min(availableCount, commentCount + 1);
  const view = new DataView(
    tableStream.buffer,
    tableStream.byteOffset,
    tableStream.byteLength,
  );
  return Array.from({ length: count }, (_, index) =>
    view.getUint32(offset + index * 4, true),
  );
}

/** 将批注 story 片段清理为列表可读的纯文本。 */
function readCommentText(
  wordDocument: Uint8Array,
  content: DocBinaryContent,
  start: number,
  end: number,
) {
  return readDocStorySegments(wordDocument, content, start, end)
    .map((segment) => segment.text)
    .join('')
    .replace(/^[\u0005\u0002]/, '')
    .replace(/[\u000d\u000a]+$/g, '')
    .replace(/\u000d/g, '\n')
    .trim();
}

/** 读取 DOC/WPS 批注作者、正文和主文档引用位置。 */
export function parseDocComments(input: {
  /** WordDocument 主流。 */
  wordDocument: Uint8Array;
  /** FIB 指定的 Table 流。 */
  tableStream: Uint8Array;
  /** 已建立 Piece Table 与格式范围的二进制内容索引。 */
  content: DocBinaryContent;
}): DocCommentParseResult {
  const references = readCommentReferences(
    input.tableStream,
    input.content.fib,
  );
  if (!references.length) return { comments: [], warnings: [] };
  const owners = readCommentOwners(input.tableStream, input.content.fib);
  const boundaries = readCommentTextBoundaries(
    input.tableStream,
    input.content.fib,
    references.length,
  );
  const fib = input.content.fib;
  const storyStart = fib.ccpText + fib.ccpFtn + fib.ccpHdr + fib.ccpMcr;
  const comments = references.map((reference, index) => {
    const localStart = boundaries[index];
    const localEnd = boundaries[index + 1];
    return {
      id: `doc-comment-${index + 1}`,
      author: owners[reference.authorIndex] || reference.initials,
      text:
        localStart !== undefined &&
        localEnd !== undefined &&
        localEnd > localStart
          ? readCommentText(
              input.wordDocument,
              input.content,
              storyStart + localStart,
              storyStart + Math.min(localEnd, fib.ccpAtn),
            )
          : '',
      charPosition: reference.charPosition,
    };
  });
  return {
    comments,
    warnings: [
      'DOC_COMMENT_RANGE_DEGRADED: DOC/WPS 二进制批注已定位到引用邻近文字，复杂批注书签范围暂未完整恢复。',
    ],
  };
}

/** 将批注标识附着到引用字符之前最近的可见正文字符。 */
export function attachDocCommentMarkers(
  segments: DocTextSegment[],
  comments: readonly DocCommentDefinition[],
) {
  if (!comments.length) return segments;
  return comments.reduce((current, comment) => {
    const preferredPosition = Math.max(0, comment.charPosition - 1);
    return current.flatMap((segment) => {
      if (
        segment.charStart === undefined ||
        segment.charEnd === undefined ||
        preferredPosition < segment.charStart ||
        preferredPosition >= segment.charEnd
      ) {
        return [segment];
      }
      const localOffset = preferredPosition - segment.charStart;
      const result: DocTextSegment[] = [];
      if (localOffset > 0) {
        result.push({
          ...segment,
          text: segment.text.slice(0, localOffset),
          charEnd: preferredPosition,
        });
      }
      result.push({
        ...segment,
        text: segment.text.slice(localOffset, localOffset + 1),
        charStart: preferredPosition,
        charEnd: preferredPosition + 1,
        bookmarkMarkers:
          localOffset === 0 ? segment.bookmarkMarkers : undefined,
        review: {
          ...segment.review,
          annotationIds: [...(segment.review?.annotationIds ?? []), comment.id],
        },
      });
      if (localOffset + 1 < segment.text.length) {
        result.push({
          ...segment,
          text: segment.text.slice(localOffset + 1),
          charStart: preferredPosition + 1,
          bookmarkMarkers: undefined,
        });
      }
      return result;
    });
  }, segments);
}

/** 在块模型中解析批注标记的稳定目标。 */
function collectCommentTargets(blocks: readonly DocBlock[]) {
  const targets = new Map<
    string,
    { blockId: string; startOffset: number; endOffset: number }
  >();
  const visitInlines = (blockId: string, inlines: readonly DocTextInline[]) => {
    let offset = 0;
    inlines.forEach((inline) => {
      if (inline.type !== 'text') return;
      inline.review?.annotationIds?.forEach((id) => {
        if (!targets.has(id)) {
          targets.set(id, {
            blockId,
            startOffset: offset,
            endOffset: offset + inline.text.length,
          });
        }
      });
      offset += inline.text.length;
    });
  };
  blocks.forEach((block) => {
    if (block.type === 'paragraph') visitInlines(block.id, block.inlines ?? []);
    else if (block.type === 'list') {
      block.items.forEach((item) => visitInlines(block.id, item.inlines ?? []));
    } else {
      block.rows.forEach((row) =>
        row.cells.forEach((cell) => visitInlines(block.id, cell.inlines ?? [])),
      );
    }
  });
  return targets;
}

/** 收集正文块中可切换显示的修订、摘要和首个定位范围。 */
function collectDocRevisionRecords(blocks: readonly DocBlock[]) {
  const collector = new WordRevisionRecordCollector();
  const visitInlines = (blockId: string, inlines: readonly DocTextInline[]) => {
    let offset = 0;
    inlines.forEach((inline) => {
      if (inline.type === 'text') {
        inline.review?.revisions?.forEach((revision) =>
          collector.add(
            revision,
            {
              kind: 'word-range',
              blockId,
              startOffset: offset,
              endOffset: offset + inline.text.length,
            },
            inline.text,
          ),
        );
        offset += inline.text.length;
      }
    });
  };
  blocks.forEach((block) => {
    const blockId = block.sourceBlockId ?? block.id;
    if (block.type === 'paragraph') {
      visitInlines(blockId, block.inlines ?? []);
    } else if (block.type === 'list') {
      block.items.forEach((item) => visitInlines(blockId, item.inlines ?? []));
    } else {
      block.rows.forEach((row) =>
        row.cells.forEach((cell) => visitInlines(blockId, cell.inlines ?? [])),
      );
    }
  });
  return collector.toArray();
}

/** 将 DOC/WPS 批注和注释摘要转换为共享审阅模型。 */
export function buildDocReviewDocument(
  blocks: readonly DocBlock[],
  comments: readonly DocCommentDefinition[],
  noteCount: number,
  warnings: readonly string[],
): WordReviewDocument | undefined {
  const revisions = collectDocRevisionRecords(blocks);
  if (!comments.length && !revisions.length && !noteCount && !warnings.length) {
    return undefined;
  }
  const targets = collectCommentTargets(blocks);
  const fallbackBlockId = blocks[0]?.id;
  const annotations = comments.flatMap((comment): OfficeAnnotation[] => {
    const target = targets.get(comment.id);
    const blockId = target?.blockId ?? fallbackBlockId;
    return blockId
      ? [
          {
            id: comment.id,
            author: comment.author,
            text: comment.text,
            target: {
              kind: 'word-range',
              blockId,
              startOffset: target?.startOffset ?? 0,
              endOffset: target?.endOffset ?? 0,
            },
          },
        ]
      : [];
  });
  return {
    annotations,
    revisions,
    revisionCount: revisions.length,
    noteCount,
    supportsRevisionModes: revisions.length > 0,
    warnings: warnings.map((warning) => {
      const separator = warning.indexOf(':');
      return separator > 0
        ? {
            code: warning.slice(0, separator),
            message: warning.slice(separator + 1).trim(),
          }
        : { code: 'DOC_REVIEW_WARNING', message: warning };
    }),
  };
}
