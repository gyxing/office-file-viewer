import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import { readXml } from '../../shared/ooxml/archive';
import type { OfficeRelationship } from '../../shared/ooxml/media';
import {
  attr,
  childByLocalName,
  descendantsByLocalName,
  parseXml,
} from '../../shared/ooxml/xml';
import type {
  SpreadsheetAutoFilter,
  SpreadsheetTable,
} from '../spreadsheet/semantics/types';
import { parseRange } from './xlsxCellFormatting';
import type {
  XlsxPackageContext,
  XlsxSheetDescriptor,
} from './XlsxPackageContext';

/** 判断 OOXML 关系是否指向当前 Sheet 的 Table 部件。 */
function isTableRelationship(relationship: OfficeRelationship) {
  return (
    /\/table$/i.test(relationship.type ?? '') ||
    /^xl\/tables\/table\d+\.xml$/i.test(relationship.target)
  );
}

/** 从 autoFilter 节点保留范围和存在条件的字段索引。 */
function parseAutoFilter(
  node: Element | null | undefined,
): SpreadsheetAutoFilter | undefined {
  const ref = attr(node, 'ref');
  if (!ref) return undefined;
  const filteredColumns = descendantsByLocalName(node, 'filterColumn').flatMap(
    (column) => {
      const value = Number(attr(column, 'colId'));
      return Number.isFinite(value) ? [value] : [];
    },
  );
  return {
    ref,
    range: parseRange(ref),
    filteredColumns: filteredColumns.length ? filteredColumns : undefined,
  };
}

/** 把单个 table.xml 转换为共享 Table 模型。 */
function parseTableXml(xml: string, fallbackId: string) {
  if (!xml) return undefined;
  const root = parseXml(xml).documentElement;
  const ref = attr(root, 'ref');
  const range = parseRange(ref);
  if (!ref || !range) return undefined;
  const style = childByLocalName(root, 'tableStyleInfo');
  const headerRowCount = Number(attr(root, 'headerRowCount') ?? 1);
  const totalsRowCount = Number(attr(root, 'totalsRowCount') ?? 0);
  const readBoolean = (value: string | undefined) =>
    value === '1' || value === 'true';
  const table: SpreadsheetTable = {
    id: attr(root, 'id') ?? fallbackId,
    name:
      attr(root, 'displayName') ?? attr(root, 'name') ?? `Table ${fallbackId}`,
    ref,
    range,
    headerRow: headerRowCount !== 0,
    totalsRow: totalsRowCount > 0,
    styleName: attr(style, 'name'),
    showRowStripes: readBoolean(attr(style, 'showRowStripes')),
    showColumnStripes: readBoolean(attr(style, 'showColumnStripes')),
    autoFilterRef: attr(childByLocalName(root, 'autoFilter'), 'ref'),
  };
  return table;
}

/** 读取物化工作表的 tableParts 和工作表级 AutoFilter。 */
export function parseMaterializedXlsxTables(input: {
  /** 当前工作表 XML 根节点。 */
  sheetNode: Element;
  /** 当前工作表关系。 */
  relationships: Record<string, OfficeRelationship>;
  /** 当前 OOXML 包条目。 */
  entries: OfficeEntryMap;
}) {
  const tableRelationshipIds = descendantsByLocalName(
    input.sheetNode,
    'tablePart',
  ).flatMap((part) => {
    const id = attr(part, 'r:id') ?? attr(part, 'id');
    return id ? [id] : [];
  });
  const tables = tableRelationshipIds.flatMap((relationshipId, index) => {
    const relationship = input.relationships[relationshipId];
    if (!relationship || !isTableRelationship(relationship)) return [];
    const table = parseTableXml(
      readXml(input.entries, relationship.target),
      `${index + 1}`,
    );
    return table ? [table] : [];
  });
  return {
    tables,
    autoFilter: parseAutoFilter(
      childByLocalName(input.sheetNode, 'autoFilter'),
    ),
  };
}

/** 按需读取当前 Sheet 关系中的 Table 部件。 */
export async function parseSourceXlsxTables(
  context: XlsxPackageContext,
  descriptor: XlsxSheetDescriptor,
  signal?: AbortSignal,
) {
  const relationships = context.relationships[descriptor.relsPath] ?? {};
  const tableRelationships =
    Object.values(relationships).filter(isTableRelationship);
  const tables = await Promise.all(
    tableRelationships.map(async (relationship, index) =>
      parseTableXml(
        await context.reader.readText(relationship.target, signal),
        `${index + 1}`,
      ),
    ),
  );
  return tables.filter((table): table is SpreadsheetTable => Boolean(table));
}

/** 流式解析 sheet autoFilter 开始标签的属性。 */
export function parseXlsxAutoFilterAttributes(
  attributes: ReadonlyMap<string, string>,
): SpreadsheetAutoFilter | undefined {
  const ref = attributes.get('ref');
  return ref ? { ref, range: parseRange(ref) } : undefined;
}
