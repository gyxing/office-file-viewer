import { attr, matchesLocalName } from '../../shared/ooxml/xml';
import type { WordRevision, WordRevisionKind } from '../word/review/types';

/** 将 DOCX 内容修订元素映射到共享 Word 修订类别。 */
export function getDocxContentRevisionKind(
  element: Element,
): WordRevisionKind | undefined {
  if (matchesLocalName(element, 'ins')) return 'insert';
  if (matchesLocalName(element, 'del')) return 'delete';
  if (matchesLocalName(element, 'moveFrom')) return 'move-from';
  if (matchesLocalName(element, 'moveTo')) return 'move-to';
  return undefined;
}

/** 读取修订元素的稳定标识、作者和日期。 */
export function parseDocxRevision(
  element: Element,
  kind: WordRevisionKind,
  fallbackId: string,
): WordRevision {
  const sourceId = attr(element, 'w:id') ?? attr(element, 'id');
  return {
    // WPS 可能在多个 rPrChange 中复用同一 w:id，位置后缀保证每一处修订独立计数。
    id: `docx-revision-${kind}-${sourceId ?? 'missing'}-${fallbackId}`,
    kind,
    author: attr(element, 'w:author') ?? attr(element, 'author'),
    createdAt: attr(element, 'w:date') ?? attr(element, 'date'),
  };
}
