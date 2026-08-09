import type { OfficeHyperlink } from './types';

/** 根据源地址语义创建可序列化的标准外部超链接。 */
export function createOfficeHyperlinkFromTarget(
  target: string,
  screenTip?: string,
): OfficeHyperlink {
  const normalized = target.trim();
  const lower = normalized.toLowerCase();
  const kind = lower.startsWith('mailto:')
    ? 'email'
    : lower.startsWith('tel:')
    ? 'phone'
    : lower.startsWith('file:') ||
      /^[a-z]:[\\/]/i.test(normalized) ||
      normalized.startsWith('\\\\')
    ? 'file'
    : 'external';
  return { kind, target: normalized, screenTip };
}
