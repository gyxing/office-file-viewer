import type { WordBookmarkTarget } from '../word/types';
import type {
  DocBinaryBookmark,
  DocBookmarkMarker,
  DocFib,
  DocTextSegment,
} from './docParseTypes';
import type { DocBlock } from './types';

/** DOC 标准书签表的安全读取结果。 */
type DocBookmarkReadResult = {
  /** 能够精确映射到主文档 story 的书签。 */
  bookmarks: DocBinaryBookmark[];
  /** 表结构异常时返回的非阻断提示。 */
  warnings: string[];
};

function tableSlice(tableStream: Uint8Array, offset: number, length: number) {
  if (!length) return new Uint8Array();
  if (offset < 0 || length < 0 || offset + length > tableStream.length) {
    throw new Error('书签表范围超出 Table 流边界');
  }
  return tableStream.slice(offset, offset + length);
}

function decodeAnsi(bytes: Uint8Array) {
  try {
    return new TextDecoder('windows-1252').decode(bytes);
  } catch {
    return Array.from(bytes, (value) => String.fromCharCode(value)).join('');
  }
}

/** 读取标准 STTB 字符串表，同时兼容扩展 Unicode 和旧式单字节结构。 */
function readBookmarkNames(data: Uint8Array) {
  if (data.length < 4) throw new Error('书签名称表缺少完整表头');
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const extended = view.getUint16(0, true) === 0xffff;
  const count = view.getUint16(extended ? 2 : 0, true);
  const extraSize = view.getUint16(extended ? 4 : 2, true);
  let offset = extended ? 6 : 4;
  if (count > 0x7ff0) throw new Error('书签名称数量超出规范上限');

  const names: string[] = [];
  for (let index = 0; index < count; index += 1) {
    const lengthFieldSize = extended ? 2 : 1;
    if (offset + lengthFieldSize > data.length) {
      throw new Error('书签名称长度字段越界');
    }
    const charCount = extended
      ? view.getUint16(offset, true)
      : data[offset] ?? 0;
    offset += lengthFieldSize;
    const byteLength = charCount * (extended ? 2 : 1);
    if (offset + byteLength + extraSize > data.length) {
      throw new Error('书签名称内容越界');
    }
    const bytes = data.slice(offset, offset + byteLength);
    names.push(
      (extended ? new TextDecoder('utf-16le').decode(bytes) : decodeAnsi(bytes))
        .replace(/\u0000/g, '')
        .trim(),
    );
    offset += byteLength + extraSize;
  }
  return names;
}

function readBookmarkRanges(
  startsData: Uint8Array,
  endsData: Uint8Array,
  names: readonly string[],
  mainStoryLength: number,
) {
  if (startsData.length < 4 || (startsData.length - 4) % 8 !== 0) {
    throw new Error('书签起点 PLC 长度不合法');
  }
  if (endsData.length < 4 || endsData.length % 4 !== 0) {
    throw new Error('书签终点 PLC 长度不合法');
  }
  const startCount = (startsData.length - 4) / 8;
  const endCount = endsData.length / 4 - 1;
  if (startCount !== names.length || startCount !== endCount) {
    throw new Error('书签名称、起点和终点表数量不一致');
  }

  const startsView = new DataView(
    startsData.buffer,
    startsData.byteOffset,
    startsData.byteLength,
  );
  const endsView = new DataView(
    endsData.buffer,
    endsData.byteOffset,
    endsData.byteLength,
  );
  const records = new Map<string, DocBinaryBookmark>();
  const metadataOffset = (startCount + 1) * 4;

  for (let index = 0; index < startCount; index += 1) {
    const name = names[index];
    const charStart = startsView.getUint32(index * 4, true);
    const endIndex = startsView.getUint16(metadataOffset + index * 4, true);
    if (!name || endIndex >= endCount) continue;
    const charEnd = endsView.getUint32(endIndex * 4, true);
    // 当前查看器只能渲染主文档 story，其他 story 的书签不能伪造目标。
    if (
      charStart > charEnd ||
      charStart >= mainStoryLength ||
      charEnd > mainStoryLength
    ) {
      continue;
    }
    records.set(name, {
      name,
      charStart,
      charEnd,
      markerId: `doc-bookmark-${index}`,
    });
  }
  return [...records.values()].sort(
    (left, right) => left.charStart - right.charStart,
  );
}

/** 按 MS-DOC 的平行 STTB/PLC 关系读取主文档标准书签。 */
export function readDocBookmarks(
  tableStream: Uint8Array,
  fib: DocFib,
): DocBookmarkReadResult {
  if (!fib.lcbSttbfBkmk && !fib.lcbPlcfBkf && !fib.lcbPlcfBkl) {
    return { bookmarks: [], warnings: [] };
  }
  try {
    const names = readBookmarkNames(
      tableSlice(tableStream, fib.fcSttbfBkmk, fib.lcbSttbfBkmk),
    );
    const bookmarks = readBookmarkRanges(
      tableSlice(tableStream, fib.fcPlcfBkf, fib.lcbPlcfBkf),
      tableSlice(tableStream, fib.fcPlcfBkl, fib.lcbPlcfBkl),
      names,
      fib.ccpText,
    );
    return { bookmarks, warnings: [] };
  } catch (error) {
    const reason = error instanceof Error ? error.message : '未知书签表错误';
    return {
      bookmarks: [],
      warnings: [`UNSUPPORTED_HYPERLINK: DOC/WPS 书签未恢复（${reason}）。`],
    };
  }
}

function markerFromBookmark(bookmark: DocBinaryBookmark): DocBookmarkMarker {
  return { name: bookmark.name, markerId: bookmark.markerId };
}

/** 在书签起点处分割文本片段，避免把链接目标粗略吸附到整段开头。 */
export function attachDocBookmarkMarkers(
  segments: readonly DocTextSegment[],
  bookmarks: readonly DocBinaryBookmark[],
) {
  if (!bookmarks.length) return [...segments];
  const sortedBookmarks = [...bookmarks].sort(
    (left, right) => left.charStart - right.charStart,
  );
  let bookmarkIndex = 0;

  return segments.flatMap((segment) => {
    if (segment.charStart === undefined || segment.charEnd === undefined) {
      return [segment];
    }
    const segmentStart = segment.charStart;
    const segmentEnd = segment.charEnd;
    const scoped: DocBinaryBookmark[] = [];
    // DOC 的零宽书签可能落在段落标记或域代码等不可见字符上，应吸附到其后的首个可见片段。
    while (
      bookmarkIndex < sortedBookmarks.length &&
      sortedBookmarks[bookmarkIndex].charStart <= segmentStart
    ) {
      scoped.push(sortedBookmarks[bookmarkIndex]);
      bookmarkIndex += 1;
    }
    while (
      bookmarkIndex < sortedBookmarks.length &&
      sortedBookmarks[bookmarkIndex].charStart < segmentEnd
    ) {
      scoped.push(sortedBookmarks[bookmarkIndex]);
      bookmarkIndex += 1;
    }
    if (!scoped.length) return [segment];

    const markersByOffset = new Map<number, DocBookmarkMarker[]>();
    scoped.forEach((bookmark) => {
      const offset = Math.max(0, bookmark.charStart - segmentStart);
      const markers = markersByOffset.get(offset) ?? [];
      markers.push(markerFromBookmark(bookmark));
      markersByOffset.set(offset, markers);
    });
    const boundaries = [0, ...markersByOffset.keys(), segment.text.length]
      .filter(
        (value, index, values) =>
          value >= 0 &&
          value <= segment.text.length &&
          values.indexOf(value) === index,
      )
      .sort((left, right) => left - right);

    return boundaries.slice(0, -1).map((start, index) => {
      const end = boundaries[index + 1];
      return {
        ...segment,
        text: segment.text.slice(start, end),
        charStart: segmentStart + start,
        charEnd: segmentStart + end,
        bookmarkMarkers: markersByOffset.get(start),
      };
    });
  });
}

function blockInlines(block: DocBlock) {
  if (block.type === 'paragraph') return block.inlines ?? [];
  if (block.type === 'table') {
    return block.rows.flatMap((row) =>
      row.cells.flatMap((cell) => cell.inlines ?? []),
    );
  }
  return block.items.flatMap((item) => item.inlines ?? []);
}

/** 返回当前块包含的书签标记，供分页索引精确定位拆分后的表格页。 */
export function docBookmarkMarkerIdsFromBlock(block: DocBlock) {
  return blockInlines(block).flatMap((inline) =>
    inline.type === 'bookmark' ? [inline.markerId] : [],
  );
}

/** 从最终块模型建立书签名称到块和零宽标记的定位表。 */
export function buildDocBookmarkTargets(blocks: readonly DocBlock[]) {
  const targets: Record<string, WordBookmarkTarget> = {};
  blocks.forEach((block) => {
    blockInlines(block).forEach((inline) => {
      if (inline.type !== 'bookmark') return;
      targets[inline.name] = {
        name: inline.name,
        targetBlockId: block.id,
        markerId: inline.markerId,
      };
    });
  });
  return targets;
}
