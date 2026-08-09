import type { OfficeHyperlink } from '../../shared/hyperlink';

function externalHyperlink(
  target: string,
  screenTip?: string,
): OfficeHyperlink {
  const normalized = target.trim();
  const lower = normalized.toLowerCase();
  return {
    kind: lower.startsWith('mailto:')
      ? 'email'
      : lower.startsWith('tel:')
      ? 'phone'
      : lower.startsWith('file:') ||
        /^[a-z]:[\\/]/i.test(normalized) ||
        normalized.startsWith('\\\\')
      ? 'file'
      : 'external',
    target: normalized,
    screenTip,
  };
}

/** 解析 DOC/WPS HYPERLINK 域中可静态确定的链接。 */
export function parseDocFieldHyperlink(
  instruction: string,
): OfficeHyperlink | undefined {
  if (!/^\s*HYPERLINK\b/i.test(instruction)) return undefined;
  const bookmark = /\\l\s+(?:"([^"]+)"|'([^']+)'|(\S+))/i.exec(instruction);
  const screenTip = /\\o\s+(?:"([^"]+)"|'([^']+)'|(\S+))/i.exec(instruction);
  const tip = screenTip?.[1] ?? screenTip?.[2] ?? screenTip?.[3];
  if (bookmark) {
    const name = bookmark[1] ?? bookmark[2] ?? bookmark[3];
    return name
      ? {
          kind: 'internal',
          target: { family: 'word', bookmark: name },
          screenTip: tip,
        }
      : undefined;
  }
  const target = /^\s*HYPERLINK\s+(?:"([^"]+)"|'([^']+)'|(\S+))/i.exec(
    instruction,
  );
  const value = target?.[1] ?? target?.[2] ?? target?.[3];
  return value ? externalHyperlink(value, tip) : undefined;
}
