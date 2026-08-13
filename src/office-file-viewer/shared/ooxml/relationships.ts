import { normalizeRelationshipTarget, type OfficeRelationship } from './media';
import { attr, descendantsByLocalName, parseXml } from './xml';

/** 返回 OPC 部件对应的关系文件路径。 */
export function getOfficePartRelationshipsPath(partPath: string) {
  const separatorIndex = partPath.lastIndexOf('/');
  const directory =
    separatorIndex >= 0 ? partPath.slice(0, separatorIndex) : '';
  const fileName =
    separatorIndex >= 0 ? partPath.slice(separatorIndex + 1) : partPath;
  return `${directory ? `${directory}/` : ''}_rels/${fileName}.rels`;
}

/** 读取 OOXML 关系文件并解析目标路径。 */
export function readRelationships(xml: string, relsPath: string) {
  const doc = parseXml(xml);
  const relationships: Record<string, OfficeRelationship> = {};
  descendantsByLocalName(doc.documentElement, 'Relationship').forEach(
    (node) => {
      const id = attr(node, 'Id');
      const target = attr(node, 'Target');
      if (!id || !target) return;
      const targetMode = attr(node, 'TargetMode');
      relationships[id] = {
        id,
        // 外部关系不是压缩包路径，规范化会破坏相对 URL 和文件地址。
        target:
          targetMode?.toLowerCase() === 'external'
            ? target
            : normalizeRelationshipTarget(relsPath, target),
        type: attr(node, 'Type'),
        targetMode,
      };
    },
  );
  return relationships;
}
