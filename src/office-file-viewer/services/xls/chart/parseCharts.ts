import type {
  SpreadsheetImage,
  SpreadsheetWarning,
} from '../../spreadsheet/types';
import { Biff8Reader, Biff8RecordCursor } from '../biff8/Biff8Reader';
import {
  BIFF8_RECORD,
  BIFF8_SUBSTREAM,
  BIFF8_VERSION,
} from '../biff8/constants';
import type { Biff8DrawingShape } from '../drawing/types';
import { XlsParseError } from '../errors';
import type {
  Biff8ChartSubstream,
  Biff8SheetDescriptor,
  Biff8Workbook,
  Biff8Worksheet,
} from '../types';
import { adaptBiff8Chart } from './adaptChart';
import { parseChartSubstream } from './parseChartSubstream';

/** 从 BoundSheet8 指向的位置读取独立 Chart 子流。 */
export function readBiff8ChartSubstream(
  stream: Uint8Array,
  descriptor: Biff8SheetDescriptor,
  endOffset: number,
): Biff8ChartSubstream {
  const cursor = new Biff8RecordCursor(
    stream,
    descriptor.streamOffset,
    endOffset,
  );
  const bof = cursor.next();
  if (!bof || bof.id !== BIFF8_RECORD.BOF || bof.data.length < 4) {
    throw new XlsParseError(
      'INVALID_RECORD_DATA',
      `图表工作表“${descriptor.name}”缺少 BOF`,
    );
  }
  const reader = new Biff8Reader(bof.data);
  if (
    reader.readUint16() !== BIFF8_VERSION ||
    reader.readUint16() !== BIFF8_SUBSTREAM.CHART
  ) {
    throw new XlsParseError(
      'INVALID_RECORD_DATA',
      `图表工作表“${descriptor.name}”不是 BIFF8 Chart 子流`,
    );
  }
  const records = [bof];
  for (let record = cursor.next(); record; record = cursor.next()) {
    records.push(record);
    if (record.id === BIFF8_RECORD.EOF) {
      return { offset: descriptor.streamOffset, records };
    }
  }
  throw new XlsParseError(
    'INVALID_RECORD_DATA',
    `图表工作表“${descriptor.name}”缺少 EOF`,
  );
}

/** 工作表内嵌图表及其锚点和预览图片。 */
export type ParsedSheetChart = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 面向用户展示的标题。 */
  title?: string;
  /** 转换后的标准图表模型。 */
  chart: ReturnType<typeof adaptBiff8Chart>['chart'];
  /** 对象在工作表或画布中的定位锚点。 */
  anchor: Biff8DrawingShape['anchor'];
  /** 匹配到的图表静态预览图片标识。 */
  previewImageId?: string;
  /** 解析时产生但不阻止继续预览的警告。 */
  warnings: SpreadsheetWarning[];
};

/** 解析工作表内嵌图表；单个图表失败不会影响同表的其他内容。 */
export function parseBiff8Charts(
  workbook: Biff8Workbook,
  descriptor: Biff8SheetDescriptor,
  substreams: Biff8ChartSubstream[],
  shapes: Biff8DrawingShape[],
  images: SpreadsheetImage[],
  sourceSheet?: Biff8Worksheet,
) {
  const chartShapes = shapes.filter(
    (shape) =>
      shape.name?.includes('图表') ||
      (!shape.blipIndex && shape.shapeType === 75),
  );
  const results: ParsedSheetChart[] = [];
  substreams.forEach((substream, chartIndex) => {
    const shape = chartShapes[chartIndex];
    const preview = shape?.shapeId
      ? images.find((image) => image.id === `xls-image-${shape.shapeId}`)
      : undefined;
    try {
      const parsed = parseChartSubstream({
        substream,
        workbook,
        sourceSheet,
        descriptor,
        chartIndex,
        anchor: shape?.anchor,
        previewImageSrc:
          typeof preview?.src === 'string' ? preview.src : undefined,
      });
      const adapted = adaptBiff8Chart(parsed);
      results.push({
        id: parsed.id,
        title: parsed.title,
        chart: adapted.chart,
        anchor: parsed.anchor,
        previewImageId: preview?.id,
        warnings: [
          ...adapted.warnings,
          ...(!shape && sourceSheet
            ? [
                {
                  code: 'AMBIGUOUS_CHART_ANCHOR',
                  message: '未找到明确的图表形状，已使用默认锚点',
                  sheetName: descriptor.name,
                },
              ]
            : []),
        ],
      });
    } catch (error) {
      results.push({
        id: `xls-chart-${descriptor.id}-${chartIndex + 1}`,
        chart: {
          type: 'unknown',
          categories: [],
          series: [],
          renderMode: 'snapshot',
          snapshotSrc:
            typeof preview?.src === 'string' ? preview.src : undefined,
          degradedFrom: '损坏的 BIFF8 图表',
        },
        anchor: shape?.anchor ?? {
          from: { row: 0, column: 0, rowFraction: 0, columnFraction: 0 },
          to: { row: 19, column: 9, rowFraction: 0, columnFraction: 0 },
        },
        previewImageId: preview?.id,
        warnings: [
          {
            code: 'INVALID_CHART',
            message: `图表解析失败：${
              error instanceof Error ? error.message : '未知错误'
            }`,
            sheetName: descriptor.name,
            offset: substream.offset,
          },
        ],
      });
    }
  });
  return results;
}
