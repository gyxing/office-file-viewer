import type { CfbStreamReader } from '../../../shared/binary/cfb';
import type { OfficeChartModel } from '../../../shared/ooxml/charts';
import { parseXlsCore } from '../../xls/parseXlsCore';
import { PPT_RECORD } from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import { readPptPersistObject } from '../readPptPersistObject';
import type { PptEditChain, PptParseContext, PptRecord } from '../types';

/** PPT 外嵌 OLE 对象原子记录的类型编号。 */
const EX_OLE_OBJ_ATOM = 0x0fc3;

/** 将输入标准化为 `normalizeEmbeddedChart` 返回的结构。 */
function normalizeEmbeddedChart(chart: OfficeChartModel) {
  if (chart.type !== 'pie' && chart.type !== 'doughnut') return chart;
  chart.series.forEach((series) => {
    // MS Graph 的“自动配色”有时被 BIFF8 表示成白色，清除后交给统一调色板逐扇区着色。
    if (
      series.color?.toLowerCase() === '#ffffff' &&
      !series.pointColors?.length
    ) {
      series.color = undefined;
    }
  });
  return chart;
}

/** 提取并汇总 `collectRecords` 返回的数据。 */
function collectRecords(
  stream: Uint8Array,
  record: PptRecord,
  type: number,
  result: PptRecord[],
) {
  const reader = new PptRecordReader(
    stream,
    record.dataOffset,
    record.endOffset,
  );
  for (const child of reader.records()) {
    if (child.type === type) result.push(child);
    if (child.version === 0x0f && child.length >= 8) {
      collectRecords(stream, child, type, result);
    }
  }
}

async function inflateOleStorage(record: PptRecord) {
  if (record.data.length < 6) throw new Error('OLE 存储记录长度不足');
  if (typeof DecompressionStream === 'undefined') {
    throw new Error('当前浏览器缺少 deflate 解压能力');
  }
  const expectedLength = new DataView(
    record.data.buffer,
    record.data.byteOffset,
    4,
  ).getUint32(0, true);
  const stream = new Blob([record.data.subarray(4)])
    .stream()
    .pipeThrough(new DecompressionStream('deflate'));
  const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
  if (expectedLength && bytes.length !== expectedLength) {
    throw new Error('OLE 解压后长度与记录声明不一致');
  }
  return bytes;
}

/** 解析文档的 MS Graph/OLE 工作簿，并复用 BIFF8 图表适配结果。 */
export async function readPptEmbeddedCharts(
  documentStream: Uint8Array,
  editChain: PptEditChain,
  context: PptParseContext,
) {
  const documentOffset = editChain.persistOffsets.get(
    editChain.documentPersistId,
  );
  if (documentOffset === undefined) return context.charts;
  const documentRecord = new PptRecordReader(
    documentStream,
    documentOffset,
    documentStream.length,
  ).readRecord();
  if (!documentRecord) return context.charts;

  const objectAtoms: PptRecord[] = [];
  collectRecords(documentStream, documentRecord, EX_OLE_OBJ_ATOM, objectAtoms);
  for (const atom of objectAtoms) {
    if (atom.data.length < 20) continue;
    const view = new DataView(
      atom.data.buffer,
      atom.data.byteOffset,
      atom.data.byteLength,
    );
    const objectId = view.getUint32(8, true);
    const persistId = view.getUint32(16, true);
    const offset = editChain.persistOffsets.get(persistId);
    if (offset === undefined) continue;

    try {
      const storageRecord = new PptRecordReader(
        documentStream,
        offset,
        documentStream.length,
      ).readRecord();
      if (!storageRecord || storageRecord.type !== PPT_RECORD.EX_OLE_OBJ_STG) {
        continue;
      }
      const bytes = await inflateOleStorage(storageRecord);
      const result = await parseXlsCore(bytes, {
        checkpoint: () => context.yieldIfNeeded(),
      });
      const chart = result.workbook.sheets
        .flatMap((sheet) => sheet.charts)
        .find((item) => item.chart);
      if (chart) {
        context.charts.set(objectId, {
          chart: normalizeEmbeddedChart(chart.chart),
          title: chart.title,
        });
      }
    } catch (error) {
      context.warnings.push({
        code: 'PPT_EMBEDDED_CHART_FALLBACK',
        message: `嵌入图表已保留静态预览：${
          error instanceof Error ? error.message : '工作簿解析失败'
        }`,
        offset,
      });
    }
    await context.yieldIfNeeded();
  }
  return context.charts;
}

/** 从已读取的根文档和按 Persist Offset 获取的 OLE 记录恢复嵌入图表。 */
export async function readPptEmbeddedChartsFromStream(
  documentRecordBytes: Uint8Array,
  documentStream: CfbStreamReader,
  editChain: PptEditChain,
  context: PptParseContext,
  signal?: AbortSignal,
) {
  const documentRecord = new PptRecordReader(documentRecordBytes).readRecord();
  if (!documentRecord) return context.charts;
  const objectAtoms: PptRecord[] = [];
  collectRecords(
    documentRecordBytes,
    documentRecord,
    EX_OLE_OBJ_ATOM,
    objectAtoms,
  );
  for (const atom of objectAtoms) {
    if (atom.data.length < 20) continue;
    const view = new DataView(
      atom.data.buffer,
      atom.data.byteOffset,
      atom.data.byteLength,
    );
    const objectId = view.getUint32(8, true);
    const persistId = view.getUint32(16, true);
    const loaded = await readPptPersistObject(
      documentStream,
      editChain,
      persistId,
      signal,
    );
    if (!loaded) continue;

    try {
      const storageRecord = new PptRecordReader(loaded.bytes).readRecord();
      if (!storageRecord || storageRecord.type !== PPT_RECORD.EX_OLE_OBJ_STG) {
        continue;
      }
      const bytes = await inflateOleStorage(storageRecord);
      const result = await parseXlsCore(bytes, {
        checkpoint: () => context.yieldIfNeeded(),
      });
      const chart = result.workbook.sheets
        .flatMap((sheet) => sheet.charts)
        .find((item) => item.chart);
      if (chart) {
        context.charts.set(objectId, {
          chart: normalizeEmbeddedChart(chart.chart),
          title: chart.title,
        });
      }
    } catch (error) {
      context.warnings.push({
        code: 'PPT_EMBEDDED_CHART_FALLBACK',
        message: `嵌入图表已保留静态预览：${
          error instanceof Error ? error.message : '工作簿解析失败'
        }`,
        offset: loaded.sourceOffset,
      });
    }
    await context.yieldIfNeeded();
  }
  return context.charts;
}
