import { emuToPx } from '../../shared/ooxml/units';
import {
  attr,
  childByLocalName,
  descendantByLocalName,
  descendantsByLocalName,
  parseXml,
  textContent,
} from '../../shared/ooxml/xml';
import { createSpreadsheetAxisIndex } from '../spreadsheet/SpreadsheetAxisIndex';
import type { SpreadsheetSheetLayout } from '../spreadsheet/SpreadsheetSource';
import type {
  SpreadsheetAnchorPoint,
  SpreadsheetImage,
} from '../spreadsheet/types';
import { createXlsxImageResource } from './createXlsxImageResource';
import type {
  XlsxPackageContext,
  XlsxSheetDescriptor,
} from './XlsxPackageContext';

function readAnchorPoint(node: Element | null): SpreadsheetAnchorPoint {
  return {
    column: Number(textContent(childByLocalName(node, 'col'))) + 1,
    columnOffset: emuToPx(
      Number(textContent(childByLocalName(node, 'colOff')) || 0),
    ),
    row: Number(textContent(childByLocalName(node, 'row'))) + 1,
    rowOffset: emuToPx(
      Number(textContent(childByLocalName(node, 'rowOff')) || 0),
    ),
  };
}

function anchorPointKey(point: SpreadsheetAnchorPoint) {
  return `${point.column}:${point.columnOffset}:${point.row}:${point.rowOffset}`;
}

/** 用锚点识别 drawing 与 OLE objectPr 对同一预览图的重复描述。 */
export function xlsxImageAnchorKey(image: SpreadsheetImage) {
  return `${anchorPointKey(image.from)}-${anchorPointKey(image.to)}`;
}

/** OLE 预览图优先于 drawing 占位图，合并时按完整锚点去重。 */
export function mergeXlsxPreviewImages(
  drawingImages: readonly SpreadsheetImage[],
  oleImages: readonly SpreadsheetImage[],
) {
  const oleAnchors = new Set(oleImages.map(xlsxImageAnchorKey));
  return [
    ...drawingImages.filter(
      (image) => !oleAnchors.has(xlsxImageAnchorKey(image)),
    ),
    ...oleImages,
  ];
}

/** 读取 worksheet/objectPr 中的嵌入对象预览图。 */
export async function loadXlsxOlePreviewImages(
  context: XlsxPackageContext,
  descriptor: XlsxSheetDescriptor,
  layout: SpreadsheetSheetLayout,
  signal?: AbortSignal,
): Promise<SpreadsheetImage[]> {
  const relationships = context.relationships[descriptor.relsPath] ?? {};
  if (
    !Object.values(relationships).some((relationship) =>
      relationship.type?.endsWith('/image'),
    )
  ) {
    return [];
  }
  const xml = await context.reader.readText(descriptor.path, signal);
  const document = parseXml(xml);
  const rowAxis = createSpreadsheetAxisIndex(
    layout.rowCount,
    layout.defaultRowHeight,
    layout.rows,
  );
  const columnAxis = createSpreadsheetAxisIndex(
    layout.columnCount,
    layout.defaultColumnWidth,
    layout.columns,
  );
  return descendantsByLocalName(document.documentElement, 'objectPr')
    .map((objectPr, index): SpreadsheetImage | undefined => {
      const relationshipId = attr(objectPr, 'r:id') ?? attr(objectPr, 'id');
      const imagePath = relationshipId
        ? relationships[relationshipId]?.target
        : undefined;
      const source = imagePath
        ? createXlsxImageResource(context, imagePath)
        : undefined;
      const anchor = descendantByLocalName(objectPr, 'anchor');
      if (!source || !anchor) return undefined;
      const from = readAnchorPoint(childByLocalName(anchor, 'from'));
      const to = readAnchorPoint(childByLocalName(anchor, 'to'));
      const x = columnAxis.offsetAt(from.column) + from.columnOffset;
      const y = rowAxis.offsetAt(from.row) + from.rowOffset;
      const right = columnAxis.offsetAt(to.column) + to.columnOffset;
      const bottom = rowAxis.offsetAt(to.row) + to.rowOffset;
      return {
        id: `${descriptor.path}-ole-preview-${index + 1}`,
        name: '嵌入对象预览',
        alt: '嵌入对象预览',
        src: source,
        from,
        to,
        x,
        y,
        width: Math.max(1, right - x),
        height: Math.max(1, bottom - y),
      };
    })
    .filter(Boolean) as SpreadsheetImage[];
}
