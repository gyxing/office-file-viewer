import type { OfficeHyperlink, OfficeInternalHyperlinkTarget } from './types';

/** 共享导航层可以安全执行的目标分类。 */
export type ClassifiedOfficeHyperlink =
  | { kind: 'internal'; target: OfficeInternalHyperlinkTarget }
  | { kind: 'new-tab'; target: string }
  | { kind: 'system'; target: string }
  | { kind: 'host-only'; target: string }
  | { kind: 'blocked'; target: string };

/** 浏览器默认导航明确禁止的可执行协议。 */
const BLOCKED_PROTOCOLS = new Set(['javascript:', 'data:', 'vbscript:']);

/** 去除会干扰协议判断、但不属于合法地址内容的控制字符。 */
function normalizeHyperlinkTarget(target: string) {
  return target.replace(/[\u0000-\u001f\u007f]/g, '').trim();
}

/** 判断 URL 是否可以作为远程文档中相对地址的可靠基准。 */
function getRemoteBaseUrl(sourceUrl: string | undefined) {
  if (!sourceUrl) return undefined;
  try {
    const baseUrl =
      typeof window === 'undefined' ? undefined : window.location.href;
    const url = baseUrl ? new URL(sourceUrl, baseUrl) : new URL(sourceUrl);
    return url.protocol === 'http:' || url.protocol === 'https:'
      ? url
      : undefined;
  } catch {
    return undefined;
  }
}

/** 将源链接归一为内部、外部、宿主接管或安全拦截结果。 */
export function classifyOfficeHyperlink(
  hyperlink: OfficeHyperlink,
  sourceUrl?: string,
): ClassifiedOfficeHyperlink {
  if (hyperlink.kind === 'internal') {
    return { kind: 'internal', target: hyperlink.target };
  }

  const target = normalizeHyperlinkTarget(hyperlink.target);
  if (!target) return { kind: 'blocked', target };
  if (hyperlink.kind === 'file' || /^[a-z]:[\\/]/i.test(target)) {
    return { kind: 'host-only', target };
  }
  if (/^\\\\/.test(target)) return { kind: 'host-only', target };

  try {
    const absolute = new URL(target);
    const protocol = absolute.protocol.toLowerCase();
    if (BLOCKED_PROTOCOLS.has(protocol)) return { kind: 'blocked', target };
    if (protocol === 'http:' || protocol === 'https:') {
      return { kind: 'new-tab', target: absolute.href };
    }
    if (protocol === 'mailto:' || protocol === 'tel:') {
      return { kind: 'system', target: absolute.href };
    }
    if (protocol === 'file:') return { kind: 'host-only', target };
    return { kind: 'blocked', target };
  } catch {
    const baseUrl = getRemoteBaseUrl(sourceUrl);
    if (!baseUrl) return { kind: 'host-only', target };
    try {
      const resolved = new URL(target, baseUrl);
      return resolved.protocol === 'http:' || resolved.protocol === 'https:'
        ? { kind: 'new-tab', target: resolved.href }
        : { kind: 'blocked', target };
    } catch {
      return { kind: 'blocked', target };
    }
  }
}
