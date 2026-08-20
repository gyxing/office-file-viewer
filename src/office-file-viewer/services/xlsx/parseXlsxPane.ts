import { attr, childByLocalName } from '../../shared/ooxml/xml';
import type { SpreadsheetPane } from '../spreadsheet/semantics/types';

/** 将 OOXML pane 属性转换为共享冻结窗格模型。 */
export function parseXlsxPaneAttributes(
  attributes: ReadonlyMap<string, string>,
): SpreadsheetPane | undefined {
  const state = attributes.get('state') as SpreadsheetPane['state'];
  const splitX = Number(attributes.get('xSplit') ?? 0);
  const splitY = Number(attributes.get('ySplit') ?? 0);
  const frozen = state === 'frozen' || state === 'frozenSplit';
  if (
    !state &&
    !splitX &&
    !splitY &&
    !attributes.get('topLeftCell') &&
    !attributes.get('activePane')
  ) {
    return undefined;
  }
  return {
    frozenRows: frozen ? Math.max(0, Math.trunc(splitY)) : 0,
    frozenColumns: frozen ? Math.max(0, Math.trunc(splitX)) : 0,
    topLeftCell: attributes.get('topLeftCell'),
    activePane: attributes.get('activePane'),
    state,
    splitX: splitX || undefined,
    splitY: splitY || undefined,
  };
}

/** 从物化工作表 DOM 中读取首个 sheetView/pane。 */
export function parseMaterializedXlsxPane(sheetNode: Element) {
  const sheetViews = childByLocalName(sheetNode, 'sheetViews');
  const sheetView = childByLocalName(sheetViews, 'sheetView');
  const pane = childByLocalName(sheetView, 'pane');
  if (!pane) return undefined;
  const attributes = new Map<string, string>();
  Array.from(pane.attributes).forEach((attribute) => {
    attributes.set(attribute.localName || attribute.name, attribute.value);
  });
  // attr 兼容自定义 DOM 没有 localName 的情况。
  ['xSplit', 'ySplit', 'topLeftCell', 'activePane', 'state'].forEach((name) => {
    const value = attr(pane, name);
    if (value !== undefined) attributes.set(name, value);
  });
  return parseXlsxPaneAttributes(attributes);
}
