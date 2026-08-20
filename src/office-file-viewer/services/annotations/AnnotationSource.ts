import type { WordRevisionRecord } from '../word/review/types';
import type { OfficeAnnotation, OfficeAnnotationSourceSnapshot } from './types';

/** 按索引范围读取批注时使用的只读数据源。 */
export interface OfficeAnnotationSource {
  /** 返回当前批注数量、能力与数据修订号。 */
  getSnapshot(): OfficeAnnotationSourceSnapshot;
  /** 订阅数据源变化，并返回取消订阅函数。 */
  subscribe(listener: () => void): () => void;
  /** 按文档顺序读取左闭右开的批注范围。 */
  getRange(
    start: number,
    end: number,
    signal?: AbortSignal,
  ): Promise<readonly OfficeAnnotation[]>;
  /** 按稳定标识查找批注索引；未知时返回 -1。 */
  findIndexById(id: string, signal?: AbortSignal): Promise<number>;
  /** 按文档顺序读取左闭右开的修订范围；非 Word 数据源可省略。 */
  getRevisionRange?(
    start: number,
    end: number,
    signal?: AbortSignal,
  ): Promise<readonly WordRevisionRecord[]>;
  /** 按稳定标识查找修订索引；非 Word 数据源可省略。 */
  findRevisionIndexById?(id: string, signal?: AbortSignal): Promise<number>;
}

/** 创建与浏览器异步任务一致的批注取消错误。 */
function createAnnotationAbortError() {
  if (typeof DOMException !== 'undefined') {
    return new DOMException('Office 批注读取已取消', 'AbortError');
  }
  const error = new Error('Office 批注读取已取消');
  error.name = 'AbortError';
  return error;
}

/** 在读取批注前统一检查取消信号。 */
function throwIfAnnotationReadAborted(signal?: AbortSignal) {
  if (signal?.aborted) throw createAnnotationAbortError();
}

/** 把调用方范围约束到当前批注数量内。 */
function normalizeRange(start: number, end: number, count: number) {
  const normalizedStart = Math.min(
    count,
    Math.max(0, Number.isFinite(start) ? Math.trunc(start) : 0),
  );
  const normalizedEnd = Math.min(
    count,
    Math.max(normalizedStart, Number.isFinite(end) ? Math.trunc(end) : count),
  );
  return { start: normalizedStart, end: normalizedEnd };
}

/** 创建适合物化解析结果使用的内存批注数据源。 */
export function createMemoryOfficeAnnotationSource(options: {
  /** 按文档顺序排列的全部批注。 */
  annotations: readonly OfficeAnnotation[];
  /** 按正文首次出现顺序排列的 Word 修订记录。 */
  revisions?: readonly WordRevisionRecord[];
  /** 当前文档包含的修订记录数量。 */
  revisionCount?: number;
  /** 当前文档包含的脚注、尾注或其他笔记数量。 */
  noteCount?: number;
  /** 当前 Word 文档是否支持三种修订投影。 */
  supportsRevisionModes?: boolean;
}): OfficeAnnotationSource {
  const annotations = [...options.annotations];
  const revisions = [...(options.revisions ?? [])];
  const indexById = new Map(
    annotations.map((annotation, index) => [annotation.id, index] as const),
  );
  const revisionIndexById = new Map(
    revisions.map((revision, index) => [revision.id, index] as const),
  );
  const snapshot: OfficeAnnotationSourceSnapshot = {
    revision: 1,
    count: annotations.length,
    revisionCount: Math.max(
      revisions.length,
      Math.max(0, Math.trunc(options.revisionCount ?? 0)),
    ),
    noteCount: Math.max(0, Math.trunc(options.noteCount ?? 0)),
    supportsRevisionModes: Boolean(options.supportsRevisionModes),
  };

  return {
    getSnapshot: () => snapshot,
    subscribe: () => () => undefined,
    async getRange(start, end, signal) {
      throwIfAnnotationReadAborted(signal);
      const range = normalizeRange(start, end, annotations.length);
      return annotations.slice(range.start, range.end);
    },
    async findIndexById(id, signal) {
      throwIfAnnotationReadAborted(signal);
      return indexById.get(id) ?? -1;
    },
    async getRevisionRange(start, end, signal) {
      throwIfAnnotationReadAborted(signal);
      const range = normalizeRange(start, end, revisions.length);
      return revisions.slice(range.start, range.end);
    },
    async findRevisionIndexById(id, signal) {
      throwIfAnnotationReadAborted(signal);
      return revisionIndexById.get(id) ?? -1;
    },
  };
}
