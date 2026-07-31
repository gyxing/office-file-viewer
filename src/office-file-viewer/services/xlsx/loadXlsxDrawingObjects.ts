import { parseOfficeChartXml } from '../../shared/ooxml/charts';
import { emuToPx } from '../../shared/ooxml/units';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  descendantByLocalName,
  parseXml,
  textContent,
} from '../../shared/ooxml/xml';
import { createSpreadsheetAxisIndex } from '../spreadsheet/SpreadsheetAxisIndex';
import type { SpreadsheetSheetLayout } from '../spreadsheet/SpreadsheetSource';
import type {
  SpreadsheetAnchorPoint,
  SpreadsheetChart,
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

function drawingRelationshipPath(path: string) {
  const parts = path.split('/');
  const fileName = parts.pop() ?? path;
  return `${parts.join('/')}/_rels/${fileName}.rels`;
}

/** 读取当前 Sheet 的 drawing，并让图片内容保持懒加载。 */
export async function loadXlsxDrawingObjects(
  context: XlsxPackageContext,
  descriptor: XlsxSheetDescriptor,
  layout: SpreadsheetSheetLayout,
  drawingRelationshipId: string | undefined,
  signal?: AbortSignal,
): Promise<{
  images: SpreadsheetImage[];
  charts: SpreadsheetChart[];
}> {
  const drawingPath = drawingRelationshipId
    ? context.relationships[descriptor.relsPath]?.[drawingRelationshipId]
        ?.target
    : undefined;
  if (!drawingPath || !context.reader.has(drawingPath)) {
    return { images: [], charts: [] };
  }
  const xml = await context.reader.readText(drawingPath, signal);
  const document = parseXml(xml);
  const drawingRels =
    context.relationships[drawingRelationshipPath(drawingPath)] ?? {};
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
  const images: SpreadsheetImage[] = [];
  const charts: SpreadsheetChart[] = [];
  const chartTasks: Promise<void>[] = [];

  childrenByLocalName(document.documentElement, 'twoCellAnchor').forEach(
    (anchorNode, index) => {
      const from = readAnchorPoint(childByLocalName(anchorNode, 'from'));
      const to = readAnchorPoint(childByLocalName(anchorNode, 'to'));
      const x = columnAxis.offsetAt(from.column) + from.columnOffset;
      const y = rowAxis.offsetAt(from.row) + from.rowOffset;
      const right = columnAxis.offsetAt(to.column) + to.columnOffset;
      const bottom = rowAxis.offsetAt(to.row) + to.rowOffset;
      const name = attr(descendantByLocalName(anchorNode, 'cNvPr'), 'name');
      const blip = descendantByLocalName(anchorNode, 'blip');
      const imageRelationshipId = attr(blip, 'r:embed') ?? attr(blip, 'embed');
      const imagePath = imageRelationshipId
        ? drawingRels[imageRelationshipId]?.target
        : undefined;
      if (imagePath && context.reader.has(imagePath)) {
        const source = createXlsxImageResource(context, imagePath);
        if (!source) return;
        images.push({
          id: `${drawingPath}-image-${index + 1}`,
          name,
          alt: name,
          src: source,
          from,
          to,
          x,
          y,
          width: Math.max(1, right - x),
          height: Math.max(1, bottom - y),
        });
      }

      const chartNode = descendantByLocalName(anchorNode, 'chart');
      const chartRelationshipId =
        attr(chartNode, 'r:id') ?? attr(chartNode, 'id');
      const chartPath = chartRelationshipId
        ? drawingRels[chartRelationshipId]?.target
        : undefined;
      if (chartPath && context.reader.has(chartPath)) {
        charts.push({
          id: `${drawingPath}-chart-${index + 1}`,
          title: name,
          chart: { type: 'bar', categories: [], series: [] },
          from,
          to,
          x,
          y,
          width: Math.max(1, right - x),
          height: Math.max(1, bottom - y),
        });
        const target = charts[charts.length - 1];
        chartTasks.push(
          context.reader.readText(chartPath, signal).then((chartXml) => {
            target.chart = parseOfficeChartXml(chartXml, context.theme);
          }),
        );
      }
    },
  );
  await Promise.all(chartTasks);
  return { images, charts };
}
