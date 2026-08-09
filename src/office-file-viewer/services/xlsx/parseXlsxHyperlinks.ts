import {
  createOfficeHyperlinkFromTarget,
  type OfficeHyperlink,
} from '../../shared/hyperlink';
import type { OfficeRelationship } from '../../shared/ooxml/media';
import { attr, descendantByLocalName } from '../../shared/ooxml/xml';
import type {
  SpreadsheetCell,
  SpreadsheetHyperlinkRange,
} from '../spreadsheet/types';
import { parseRange } from './xlsxCellFormatting';

/** 将电子表格外部目标转换为共享链接模型。 */
export function createSpreadsheetExternalHyperlink(
  target: string,
  screenTip?: string,
): OfficeHyperlink {
  const normalized = target.trim();
  if (normalized.startsWith('#')) {
    return internalHyperlink(normalized.slice(1), screenTip);
  }
  return createOfficeHyperlinkFromTarget(normalized, screenTip);
}

function unquoteSheetName(value: string) {
  const trimmed = value.trim();
  return trimmed.startsWith("'") && trimmed.endsWith("'")
    ? trimmed.slice(1, -1).replace(/''/g, "'")
    : trimmed;
}

/** 将 Excel location 或 HYPERLINK 的井号目标转换为内部导航模型。 */
export function internalHyperlink(
  location: string,
  screenTip?: string,
): OfficeHyperlink {
  const normalized = location.replace(/^#/, '').trim();
  const separator = normalized.lastIndexOf('!');
  if (separator >= 0) {
    return {
      kind: 'internal',
      target: {
        family: 'spreadsheet',
        sheetName: unquoteSheetName(normalized.slice(0, separator)),
        cellRef: normalized.slice(separator + 1).replace(/\$/g, ''),
      },
      screenTip,
    };
  }
  const absoluteFree = normalized.replace(/\$/g, '');
  return {
    kind: 'internal',
    target: /^[A-Za-z]{1,3}\d+(?::[A-Za-z]{1,3}\d+)?$/.test(absoluteFree)
      ? { family: 'spreadsheet', cellRef: absoluteFree }
      : { family: 'spreadsheet', definedName: normalized },
    screenTip,
  };
}

/** 解析 worksheet hyperlink 节点并保留稀疏范围。 */
export function parseXlsxHyperlink(
  attributes: ReadonlyMap<string, string>,
  relationships: Record<string, OfficeRelationship>,
): SpreadsheetHyperlinkRange | undefined {
  const ref = attributes.get('ref');
  const range = parseRange(ref);
  if (!ref || !range) return undefined;
  const tooltip = attributes.get('tooltip');
  const location = attributes.get('location');
  const relationshipId = attributes.get('r:id') ?? attributes.get('id');
  const relationship = relationshipId
    ? relationships[relationshipId]
    : undefined;
  const hyperlink = location
    ? internalHyperlink(location, tooltip)
    : relationship?.target
    ? createSpreadsheetExternalHyperlink(relationship.target, tooltip)
    : undefined;
  return hyperlink ? { ref, ...range, hyperlink } : undefined;
}

function readFormulaStringArguments(formula: string) {
  const match =
    /^\s*=?\s*HYPERLINK\s*\(\s*"((?:[^"]|"")*)"\s*(?:[,;]\s*"((?:[^"]|"")*)"\s*)?\)\s*$/i.exec(
      formula,
    );
  if (!match) return undefined;
  return {
    target: match[1].replace(/""/g, '"'),
    display: match[2]?.replace(/""/g, '"'),
  };
}

/** 仅解析两个参数都可静态确定的 HYPERLINK 公式。 */
export function applyStaticXlsxFormulaHyperlink(cell: SpreadsheetCell) {
  if (!cell.formula) return cell;
  const parsed = readFormulaStringArguments(cell.formula);
  if (!parsed) return cell;
  return {
    ...cell,
    value: parsed.display ?? cell.value,
    hyperlink: parsed.target.startsWith('#')
      ? internalHyperlink(parsed.target)
      : createSpreadsheetExternalHyperlink(parsed.target),
  };
}

/** 解析 DrawingML 图片或形状非可视属性中的点击链接。 */
export function parseXlsxDrawingHyperlink(
  node: Element,
  relationships: Record<string, OfficeRelationship>,
) {
  const click = descendantByLocalName(node, 'hlinkClick');
  if (!click) return undefined;
  const relationshipId = attr(click, 'r:id') ?? attr(click, 'id');
  const relationship = relationshipId
    ? relationships[relationshipId]
    : undefined;
  const tooltip = attr(click, 'tooltip');
  return relationship?.target
    ? createSpreadsheetExternalHyperlink(relationship.target, tooltip)
    : undefined;
}

/** 把范围链接仅附加到当前物化窗口中的单元格。 */
export function applyXlsxHyperlinkRanges(
  cells: Map<string, SpreadsheetCell>,
  hyperlinks: readonly SpreadsheetHyperlinkRange[],
) {
  hyperlinks.forEach((range) => {
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (
        let column = range.startColumn;
        column <= range.endColumn;
        column += 1
      ) {
        const ref = cellReference(row, column);
        const cell = cells.get(ref);
        if (cell) cell.hyperlink = range.hyperlink;
      }
    }
  });
}

function cellReference(row: number, column: number) {
  let current = column;
  let label = '';
  while (current > 0) {
    current -= 1;
    label = String.fromCharCode(65 + (current % 26)) + label;
    current = Math.floor(current / 26);
  }
  return `${label}${row}`;
}
