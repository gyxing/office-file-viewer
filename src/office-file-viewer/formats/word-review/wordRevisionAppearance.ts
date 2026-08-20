import type { CSSProperties } from 'react';
import type { WordRevision } from '../../services/word/review/types';

/** Word 按作者区分修订时使用的高对比色板。 */
const WORD_REVISION_COLORS = [
  { color: '#c2410c', tint: '#fff7ed' },
  { color: '#1d4ed8', tint: '#eff6ff' },
  { color: '#7e22ce', tint: '#faf5ff' },
  { color: '#047857', tint: '#ecfdf5' },
  { color: '#be123c', tint: '#fff1f2' },
  { color: '#0e7490', tint: '#ecfeff' },
] as const;

/** 以作者为主、修订标识为回退生成稳定色板索引。 */
function hashRevisionOwner(value: string) {
  let hash = 0;
  for (let index = 0; index < value.length; index += 1) {
    hash = (hash * 31 + value.charCodeAt(index)) >>> 0;
  }
  return hash;
}

/** 返回同一作者在正文、列表和气泡中共用的颜色。 */
export function getWordRevisionAppearance(revision: WordRevision) {
  const key = revision.author?.trim() || revision.id;
  return WORD_REVISION_COLORS[
    hashRevisionOwner(key) % WORD_REVISION_COLORS.length
  ];
}

/** 把作者颜色写入修订组件共用的 CSS 变量。 */
export function getWordRevisionCssVariables(
  revision: WordRevision,
): CSSProperties {
  const appearance = getWordRevisionAppearance(revision);
  return {
    '--office-file-revision-color': appearance.color,
    '--office-file-revision-tint': appearance.tint,
  } as CSSProperties;
}
