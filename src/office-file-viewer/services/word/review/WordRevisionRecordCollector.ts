import type { OfficeWordAnnotationTarget } from '../../annotations/types';
import type { WordRevision, WordRevisionRecord } from './types';

/** 修订列表只保留短摘要，避免大文件重复保存整段正文。 */
const WORD_REVISION_EXCERPT_LIMIT = 240;

/** 合并同一修订跨运行产生的文字片段。 */
function appendExcerpt(current: string, fragment: string) {
  const text = fragment.replace(/\s+/g, ' ').trim();
  if (!text || current.length >= WORD_REVISION_EXCERPT_LIMIT) return current;
  const separator = current ? ' ' : '';
  const next = `${current}${separator}${text}`;
  return next.length <= WORD_REVISION_EXCERPT_LIMIT
    ? next
    : `${next.slice(0, WORD_REVISION_EXCERPT_LIMIT - 1)}…`;
}

/** 在解析过程中按首次出现顺序聚合同一修订的定位与摘要。 */
export class WordRevisionRecordCollector {
  private readonly records = new Map<string, WordRevisionRecord>();

  /** 登记一段属于指定修订的正文。 */
  add(
    revision: WordRevision,
    target: OfficeWordAnnotationTarget,
    fragment = '',
  ) {
    const current = this.records.get(revision.id);
    if (current) {
      const excerpt = appendExcerpt(current.excerpt, fragment);
      if (excerpt !== current.excerpt) {
        this.records.set(revision.id, { ...current, excerpt });
      }
      return;
    }
    this.records.set(revision.id, {
      ...revision,
      excerpt: appendExcerpt('', fragment),
      target,
    });
  }

  /** 返回与源正文顺序一致的只读修订记录。 */
  toArray() {
    return [...this.records.values()];
  }
}
