import type { OfficeTheme } from '../../shared/ooxml/theme';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  descendantsByLocalName,
  textContent,
} from '../../shared/ooxml/xml';
import type {
  SpreadsheetConditionalFormattingRule,
  SpreadsheetConditionalValue,
} from '../spreadsheet/semantics/types';
import {
  parseRange,
  parseXlsxColor,
  type StyleBook,
} from './xlsxCellFormatting';

/** 流式解析期间可追加阈值和颜色的条件格式规则。 */
export type MutableXlsxConditionalRule =
  SpreadsheetConditionalFormattingRule & {
    /** 流式读取中持续追加的阈值或公式。 */
    values: SpreadsheetConditionalValue[];
    /** 流式读取中持续追加的色阶颜色。 */
    colors: string[];
  };

/** 将 Excel 条件格式类型限制到当前模型支持的联合。 */
function normalizeRuleType(
  value: string | undefined,
): SpreadsheetConditionalFormattingRule['type'] {
  return value === 'cellIs' ||
    value === 'colorScale' ||
    value === 'dataBar' ||
    value === 'iconSet' ||
    value === 'duplicateValues' ||
    value === 'uniqueValues' ||
    value === 'top10' ||
    value === 'aboveAverage' ||
    value === 'expression'
    ? value
    : 'unsupported';
}

/** 将 sqref 中的多个 A1 范围转换为标准范围数组。 */
export function parseConditionalRanges(value: string | undefined) {
  return (value ?? '')
    .trim()
    .split(/\s+/)
    .flatMap((ref) => {
      const range = parseRange(ref);
      return range ? [range] : [];
    });
}

/** 从 cfRule 属性建立可由 DOM 或流式读取继续补全的规则。 */
export function createXlsxConditionalRule(
  attributes: ReadonlyMap<string, string>,
  ranges: ReturnType<typeof parseConditionalRanges>,
  styleBook: StyleBook,
  fallbackId: string,
): MutableXlsxConditionalRule {
  const dxfId = Number(attributes.get('dxfId'));
  const type = normalizeRuleType(attributes.get('type'));
  return {
    id: attributes.get('id') ?? fallbackId,
    type,
    operator: attributes.get('operator'),
    priority: Math.max(1, Number(attributes.get('priority') ?? 1)),
    stopIfTrue:
      attributes.get('stopIfTrue') === '1' ||
      attributes.get('stopIfTrue') === 'true',
    ranges,
    style: Number.isInteger(dxfId)
      ? styleBook.differentialStyles[dxfId]
      : undefined,
    rank: Number.isFinite(Number(attributes.get('rank')))
      ? Number(attributes.get('rank'))
      : undefined,
    belowAverage:
      attributes.get('aboveAverage') === '0' ||
      attributes.get('aboveAverage') === 'false',
    values: [],
    colors: [],
  };
}

/** 将 DOM 节点属性复制为不依赖 DOM 的只读映射。 */
function attributesFromElement(element: Element) {
  const attributes = new Map<string, string>();
  Array.from(element.attributes).forEach((attribute) =>
    attributes.set(attribute.localName || attribute.name, attribute.value),
  );
  return attributes;
}

/** 解析物化工作表中的全部条件格式规则。 */
export function parseMaterializedXlsxConditionalFormatting(
  sheetNode: Element,
  styleBook: StyleBook,
  theme: OfficeTheme,
) {
  const rules: SpreadsheetConditionalFormattingRule[] = [];
  descendantsByLocalName(sheetNode, 'conditionalFormatting').forEach(
    (container) => {
      const ranges = parseConditionalRanges(attr(container, 'sqref'));
      childrenByLocalName(container, 'cfRule').forEach((ruleNode, index) => {
        const rule = createXlsxConditionalRule(
          attributesFromElement(ruleNode),
          ranges,
          styleBook,
          `xlsx-cf-${rules.length + index + 1}`,
        );
        descendantsByLocalName(ruleNode, 'formula').forEach((formula) => {
          rule.values.push({ type: 'formula', value: textContent(formula) });
        });
        descendantsByLocalName(ruleNode, 'cfvo').forEach((value) => {
          rule.values.push({
            type: attr(value, 'type') ?? 'num',
            value: attr(value, 'val'),
          });
        });
        descendantsByLocalName(ruleNode, 'color').forEach((color) => {
          const resolved = parseXlsxColor(color, theme);
          if (resolved) rule.colors.push(resolved);
        });
        const dataBar = childByLocalName(ruleNode, 'dataBar');
        rule.dataBarColor = parseXlsxColor(
          childByLocalName(dataBar, 'color'),
          theme,
        );
        const iconSet = childByLocalName(ruleNode, 'iconSet');
        rule.iconSet = attr(iconSet, 'iconSet');
        rules.push(rule);
      });
    },
  );
  return rules.sort((left, right) => left.priority - right.priority);
}
