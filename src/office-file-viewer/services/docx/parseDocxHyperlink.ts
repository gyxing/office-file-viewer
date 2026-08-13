import { createOfficeHyperlinkFromTarget } from '../../shared/hyperlink/createOfficeHyperlink';
import type { OfficeHyperlink } from '../../shared/hyperlink/types';
import type { OfficeRelationship } from '../../shared/ooxml/media';
import { attr, descendantByLocalName } from '../../shared/ooxml/xml';

export { createOfficeHyperlinkFromTarget } from '../../shared/hyperlink/createOfficeHyperlink';

/** 将 Word 书签名称转换为内部导航链接。 */
export function createDocxBookmarkHyperlink(
  bookmark: string,
  screenTip?: string,
): OfficeHyperlink {
  return {
    kind: 'internal',
    target: { family: 'word', bookmark },
    screenTip,
  };
}

/** 解析 DOCX 的显式 w:hyperlink 节点。 */
export function parseDocxHyperlinkElement(
  node: Element,
  relationships: Record<string, OfficeRelationship>,
): OfficeHyperlink | undefined {
  const screenTip = attr(node, 'w:tooltip') ?? attr(node, 'tooltip');
  const anchor = attr(node, 'w:anchor') ?? attr(node, 'anchor');
  if (anchor) return createDocxBookmarkHyperlink(anchor, screenTip);
  const relationshipId = attr(node, 'r:id') ?? attr(node, 'id');
  const relationship = relationshipId
    ? relationships[relationshipId]
    : undefined;
  return relationship?.target
    ? createOfficeHyperlinkFromTarget(relationship.target, screenTip)
    : undefined;
}

/** 解析 Word HYPERLINK 域指令中可静态确定的地址、书签和 ScreenTip。 */
export function parseDocxFieldHyperlink(
  instruction: string,
): OfficeHyperlink | undefined {
  if (!/^\s*HYPERLINK\b/i.test(instruction)) return undefined;
  const bookmark = /\\l\s+(?:"([^"]+)"|'([^']+)'|(\S+))/i.exec(instruction);
  const screenTip = /\\o\s+(?:"([^"]+)"|'([^']+)'|(\S+))/i.exec(instruction);
  const tip = screenTip?.[1] ?? screenTip?.[2] ?? screenTip?.[3];
  if (bookmark) {
    const name = bookmark[1] ?? bookmark[2] ?? bookmark[3];
    return name ? createDocxBookmarkHyperlink(name, tip) : undefined;
  }
  const target = /^\s*HYPERLINK\s+(?:"([^"]+)"|'([^']+)'|(\S+))/i.exec(
    instruction,
  );
  const value = target?.[1] ?? target?.[2] ?? target?.[3];
  return value ? createOfficeHyperlinkFromTarget(value, tip) : undefined;
}

/** 读取 DrawingML 或 VML 对象自身声明的点击链接。 */
export function parseDocxDrawingHyperlink(
  node: Element,
  relationships: Record<string, OfficeRelationship>,
): OfficeHyperlink | undefined {
  const directTarget =
    attr(node, 'href') ?? attr(node, 'o:href') ?? attr(node, 'w:href');
  if (directTarget) return createOfficeHyperlinkFromTarget(directTarget);
  const hyperlink = descendantByLocalName(node, 'hlinkClick');
  if (!hyperlink) return undefined;
  const screenTip = attr(hyperlink, 'tooltip');
  const relationshipId = attr(hyperlink, 'r:id') ?? attr(hyperlink, 'id');
  const relationship = relationshipId
    ? relationships[relationshipId]
    : undefined;
  return relationship?.target
    ? createOfficeHyperlinkFromTarget(relationship.target, screenTip)
    : undefined;
}
