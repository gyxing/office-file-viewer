import { buildDocBlocksFromSegments } from './buildDocBlocks';
import type { DocBinaryContent, DocTextSegment } from './docParseTypes';

import { readDocStorySegments } from './readDocBinaryContent';
import type { DocBlock, DocNote, DocNoteReference, DocNotes } from './types';

/** DOC 主文档中的脚注或尾注引用位置。 */
export type DocNoteReferencePosition = DocNoteReference & {
  /** 引用字符在主文档 story 中的字符位置。 */
  charPosition: number;
};

/** DOC 注释解析结果。 */
export type DocNoteParseResult = {
  /** 可按正文引用呈现的脚注和尾注正文。 */
  notes?: DocNotes;
  /** 主文档中全部脚注和尾注引用位置。 */
  references: DocNoteReferencePosition[];
  /** 损坏或暂不支持的注释结构警告。 */
  warnings: string[];
};

/** 从 PLCF 安全读取 32 位字符位置。 */
function readCpArray(tableStream: Uint8Array, offset: number, count: number) {
  if (offset < 0 || count <= 0 || offset + count * 4 > tableStream.length) {
    return [];
  }
  const view = new DataView(
    tableStream.buffer,
    tableStream.byteOffset,
    tableStream.byteLength,
  );
  return Array.from({ length: count }, (_, index) =>
    view.getUint32(offset + index * 4, true),
  );
}

/** 读取引用 PLCF 中的主文档 CP 和自动编号标志。 */
function readReferencePlc(
  tableStream: Uint8Array,
  offset: number,
  length: number,
) {
  if (!offset || length < 10 || offset + length > tableStream.length) return [];
  const count = Math.floor((length - 4) / 6);
  if (!count) return [];
  const cps = readCpArray(tableStream, offset, count + 1);
  if (cps.length !== count + 1) return [];
  const dataOffset = offset + (count + 1) * 4;
  const view = new DataView(
    tableStream.buffer,
    tableStream.byteOffset,
    tableStream.byteLength,
  );
  return cps.slice(0, count).map((charPosition, index) => ({
    charPosition,
    automatic: view.getUint16(dataOffset + index * 2, true) !== 0,
  }));
}

/** 读取仅包含 CP 的注释正文边界 PLCF。 */
function readTextBoundaries(
  tableStream: Uint8Array,
  offset: number,
  length: number,
  expectedCount: number,
) {
  if (!offset || length < 8 || offset + length > tableStream.length) return [];
  const availableCount = Math.floor(length / 4);
  return readCpArray(
    tableStream,
    offset,
    Math.min(availableCount, expectedCount + 1),
  );
}

/** 清除注释 story 自带的引用符和末尾段落标记。 */
function trimNoteSegments(segments: DocTextSegment[]) {
  const result = segments.map((segment) => ({ ...segment }));
  const first = result.find((segment) => segment.text.length);
  if (first) first.text = first.text.replace(/^[\u0002\u0005]/, '');
  for (let index = result.length - 1; index >= 0; index -= 1) {
    if (!result[index].text.length) continue;
    result[index].text = result[index].text.replace(/[\u000d\u000a]+$/, '');
    break;
  }
  return result.filter((segment) => segment.text.length);
}

/** 为每条注释生成全局唯一的块、行和单元格标识。 */
function prefixNoteBlockIds(block: DocBlock, prefix: string): DocBlock {
  if (block.type === 'paragraph')
    return { ...block, id: `${prefix}-${block.id}` };
  if (block.type === 'list') {
    return {
      ...block,
      id: `${prefix}-${block.id}`,
      items: block.items.map((item) => ({
        ...item,
        id: `${prefix}-${item.id}`,
      })),
    };
  }
  return {
    ...block,
    id: `${prefix}-${block.id}`,
    rows: block.rows.map((row) => ({
      ...row,
      id: `${prefix}-${row.id}`,
      cells: row.cells.map((cell) => ({
        ...cell,
        id: `${prefix}-${cell.id}`,
      })),
    })),
  };
}

/** 解析单类脚注或尾注的引用、正文和降级警告。 */
async function parseNoteKind(
  kind: 'footnote' | 'endnote',
  wordDocument: Uint8Array,
  tableStream: Uint8Array,
  content: DocBinaryContent,
  checkpoint: () => Promise<void>,
) {
  const fib = content.fib;
  const references = readReferencePlc(
    tableStream,
    kind === 'footnote' ? fib.fcPlcffndRef : fib.fcPlcfendRef,
    kind === 'footnote' ? fib.lcbPlcffndRef : fib.lcbPlcfendRef,
  );
  const boundaries = readTextBoundaries(
    tableStream,
    kind === 'footnote' ? fib.fcPlcffndTxt : fib.fcPlcfendTxt,
    kind === 'footnote' ? fib.lcbPlcffndTxt : fib.lcbPlcfendTxt,
    references.length,
  );
  const storyLength = kind === 'footnote' ? fib.ccpFtn : fib.ccpEdn;
  const storyStart =
    kind === 'footnote'
      ? fib.ccpText
      : fib.ccpText + fib.ccpFtn + fib.ccpHdr + fib.ccpMcr + fib.ccpAtn;
  const warnings: string[] = [];
  if (!references.length && !storyLength) {
    return { references: [], notes: [], warnings };
  }
  if (boundaries.length < references.length + 1) {
    warnings.push(
      `DOC_${
        kind === 'footnote' ? 'FOOTNOTE' : 'ENDNOTE'
      }_RANGE_UNSUPPORTED: 注释正文边界不完整，已保留最终正文。`,
    );
  }

  const notes: DocNote[] = [];
  const positionedReferences: DocNoteReferencePosition[] = [];
  for (let index = 0; index < references.length; index += 1) {
    await checkpoint();
    const reference = references[index];
    const noteId = `${index + 1}`;
    const referenceCharacter = readDocStorySegments(
      wordDocument,
      content,
      reference.charPosition,
      reference.charPosition + 1,
    )
      .map((segment) => segment.text)
      .join('');
    const label = reference.automatic
      ? `${index + 1}`
      : referenceCharacter.replace(/[\u0000-\u001f]/g, '') || `${index + 1}`;
    positionedReferences.push({
      noteId,
      noteKind: kind,
      label,
      charPosition: reference.charPosition,
    });
    const localStart = boundaries[index];
    const localEnd = boundaries[index + 1];
    if (
      localStart === undefined ||
      localEnd === undefined ||
      localStart < 0 ||
      localEnd <= localStart ||
      localEnd > storyLength
    ) {
      continue;
    }
    const segments = trimNoteSegments(
      readDocStorySegments(
        wordDocument,
        content,
        storyStart + localStart,
        storyStart + localEnd,
      ),
    );
    const blocks = await buildDocBlocksFromSegments(segments, [], {
      checkpoint,
    });
    notes.push({
      noteId,
      noteKind: kind,
      label,
      blocks: blocks.map((block) =>
        prefixNoteBlockIds(block, `doc-${kind}-${noteId}`),
      ),
    });
  }
  return { references: positionedReferences, notes, warnings };
}

/** 读取 DOC/WPS 的脚注、尾注及主文档引用。 */
export async function parseDocNotes(input: {
  /** WordDocument 主流。 */
  wordDocument: Uint8Array;
  /** FIB 指定的 Table 流。 */
  tableStream: Uint8Array;
  /** 已建立 Piece Table 与格式范围的二进制内容索引。 */
  content: DocBinaryContent;
  /** 长任务检查点。 */
  checkpoint(): Promise<void>;
}): Promise<DocNoteParseResult> {
  const footnotes = await parseNoteKind(
    'footnote',
    input.wordDocument,
    input.tableStream,
    input.content,
    input.checkpoint,
  );
  const endnotes = await parseNoteKind(
    'endnote',
    input.wordDocument,
    input.tableStream,
    input.content,
    input.checkpoint,
  );
  return {
    notes:
      footnotes.notes.length || endnotes.notes.length
        ? { footnotes: footnotes.notes, endnotes: endnotes.notes }
        : undefined,
    references: [...footnotes.references, ...endnotes.references].sort(
      (left, right) => left.charPosition - right.charPosition,
    ),
    warnings: [...footnotes.warnings, ...endnotes.warnings],
  };
}

/** 将主文档引用字符拆成独立片段，供块构建器生成可交互上标。 */
export function attachDocNoteReferences(
  segments: DocTextSegment[],
  references: readonly DocNoteReferencePosition[],
) {
  if (!references.length) return segments;
  return segments.flatMap((segment) => {
    if (segment.charStart === undefined || segment.charEnd === undefined) {
      return [segment];
    }
    const scoped = references.filter(
      (reference) =>
        reference.charPosition >= segment.charStart! &&
        reference.charPosition < segment.charEnd!,
    );
    if (!scoped.length) return [segment];
    const result: DocTextSegment[] = [];
    let cursor = 0;
    scoped.forEach((reference) => {
      const localOffset = reference.charPosition - segment.charStart!;
      if (localOffset > cursor) {
        result.push({
          ...segment,
          text: segment.text.slice(cursor, localOffset),
          charStart: segment.charStart! + cursor,
          charEnd: reference.charPosition,
          bookmarkMarkers: cursor === 0 ? segment.bookmarkMarkers : undefined,
        });
      }
      result.push({
        ...segment,
        text: segment.text.slice(localOffset, localOffset + 1),
        charStart: reference.charPosition,
        charEnd: reference.charPosition + 1,
        bookmarkMarkers:
          localOffset === 0 ? segment.bookmarkMarkers : undefined,
        noteReference: reference,
      });
      cursor = localOffset + 1;
    });
    if (cursor < segment.text.length) {
      result.push({
        ...segment,
        text: segment.text.slice(cursor),
        charStart: segment.charStart + cursor,
        bookmarkMarkers: cursor === 0 ? segment.bookmarkMarkers : undefined,
      });
    }
    return result;
  });
}
