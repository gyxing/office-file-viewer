import { parseOfficeChartXml } from '../../shared/ooxml/charts';
import { emuToPx } from '../../shared/ooxml/units';
import {
  attr,
  childByLocalName,
  descendantByLocalName,
  parseXml,
  textContent,
} from '../../shared/ooxml/xml';
import {
  createSpreadsheetAxisIndex,
  type SpreadsheetAxisIndex,
} from '../spreadsheet/SpreadsheetAxisIndex';
import type { SpreadsheetSheetLayout } from '../spreadsheet/SpreadsheetSource';
import type {
  SpreadsheetAnchorPoint,
  SpreadsheetChart,
  SpreadsheetImage,
} from '../spreadsheet/types';
import { createXlsxImageResource } from './createXlsxImageResource';
import { parseXlsxDrawingHyperlink } from './parseXlsxHyperlinks';
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

/** 按源顺序读取 XLSX 支持的双单元格和单单元格绘图锚点。 */
function readDrawingAnchorNodes(root: Element) {
  return Array.from(root.children).filter(
    (node) =>
      node.localName === 'twoCellAnchor' || node.localName === 'oneCellAnchor',
  );
}

function readDrawingExtent(anchorNode: Element) {
  const extent = childByLocalName(anchorNode, 'ext');
  if (!extent) return undefined;
  return {
    width: emuToPx(Number(attr(extent, 'cx')) || 0),
    height: emuToPx(Number(attr(extent, 'cy')) || 0),
  };
}

function anchorPointAtAxes(
  x: number,
  y: number,
  rowAxis: SpreadsheetAxisIndex,
  columnAxis: SpreadsheetAxisIndex,
): SpreadsheetAnchorPoint {
  const column = columnAxis.findIndexAtOffset(x);
  const row = rowAxis.findIndexAtOffset(y);
  return {
    column,
    columnOffset: Math.max(0, x - columnAxis.offsetAt(column)),
    row,
    rowOffset: Math.max(0, y - rowAxis.offsetAt(row)),
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
  /** 当前工作表中的浮动图片。 */
  images: SpreadsheetImage[];
  /** 当前工作表中的浮动图表。 */
  charts: SpreadsheetChart[];
  /** 浮动对象锚点覆盖到的最大行数。 */
  rowCount: number;
  /** 浮动对象锚点覆盖到的最大列数。 */
  columnCount: number;
}> {
  const drawingPath = drawingRelationshipId
    ? context.relationships[descriptor.relsPath]?.[drawingRelationshipId]
        ?.target
    : undefined;
  if (!drawingPath || !context.reader.has(drawingPath)) {
    return {
      images: [],
      charts: [],
      rowCount: 1,
      columnCount: 1,
    };
  }
  const xml = await context.reader.readText(drawingPath, signal);
  const document = parseXml(xml);
  const drawingRels =
    context.relationships[drawingRelationshipPath(drawingPath)] ?? {};
  const rawAnchors = readDrawingAnchorNodes(document.documentElement).map(
    (anchorNode) => ({
      anchorNode,
      from: readAnchorPoint(childByLocalName(anchorNode, 'from')),
      to: childByLocalName(anchorNode, 'to')
        ? readAnchorPoint(childByLocalName(anchorNode, 'to'))
        : undefined,
      extent: readDrawingExtent(anchorNode),
    }),
  );
  const estimatedRowCount = Math.max(
    1,
    ...rawAnchors.map(({ from, to, extent }) =>
      to
        ? Math.max(from.row, to.row)
        : from.row +
          Math.ceil(
            (from.rowOffset + (extent?.height ?? 0)) / layout.defaultRowHeight,
          ),
    ),
  );
  const estimatedColumnCount = Math.max(
    1,
    ...rawAnchors.map(({ from, to, extent }) =>
      to
        ? Math.max(from.column, to.column)
        : from.column +
          Math.ceil(
            (from.columnOffset + (extent?.width ?? 0)) /
              layout.defaultColumnWidth,
          ),
    ),
  );
  const rowAxis = createSpreadsheetAxisIndex(
    Math.max(layout.rowCount, estimatedRowCount),
    layout.defaultRowHeight,
    layout.rows,
  );
  const columnAxis = createSpreadsheetAxisIndex(
    Math.max(layout.columnCount, estimatedColumnCount),
    layout.defaultColumnWidth,
    layout.columns,
  );
  const anchors = rawAnchors.flatMap(({ anchorNode, from, to, extent }) => {
    const x = columnAxis.offsetAt(from.column) + from.columnOffset;
    const y = rowAxis.offsetAt(from.row) + from.rowOffset;
    if (to) return [{ anchorNode, from, to, x, y }];
    if (!extent) return [];
    return [
      {
        anchorNode,
        from,
        to: anchorPointAtAxes(
          x + extent.width,
          y + extent.height,
          rowAxis,
          columnAxis,
        ),
        x,
        y,
      },
    ];
  });
  const rowCount = Math.max(1, ...anchors.map(({ to }) => to.row));
  const columnCount = Math.max(1, ...anchors.map(({ to }) => to.column));
  const images: SpreadsheetImage[] = [];
  const charts: SpreadsheetChart[] = [];
  const chartTasks: Promise<void>[] = [];

  anchors.forEach(({ anchorNode, from, to, x, y }, index) => {
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
        hyperlink: parseXlsxDrawingHyperlink(anchorNode, drawingRels),
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
  });
  await Promise.all(chartTasks);
  return { images, charts, rowCount, columnCount };
}
