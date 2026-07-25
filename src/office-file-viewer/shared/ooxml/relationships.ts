import { normalizeRelationshipTarget, type OfficeRelationship } from './media';
import { attr, descendantsByLocalName, parseXml } from './xml';

/** 读取 `readRelationships` 所需的源数据，供 OOXML 公共解析使用。 */
export function readRelationships(xml: string, relsPath: string) {
  const doc = parseXml(xml);
  const relationships: Record<string, OfficeRelationship> = {};
  descendantsByLocalName(doc.documentElement, 'Relationship').forEach(
    (node) => {
      const id = attr(node, 'Id');
      const target = attr(node, 'Target');
      if (!id || !target) return;
      relationships[id] = {
        id,
        target: normalizeRelationshipTarget(relsPath, target),
        type: attr(node, 'Type'),
      };
    },
  );
  return relationships;
}
