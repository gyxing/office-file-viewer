import { CfbParseError, parseCfb } from '../../shared/binary/cfb';
import { createResourceReference } from '../parsing/assembly/resourceReferences';
import type { PortableResource } from '../parsing/protocol/messages';
import type { ParseProgress } from '../parsing/types';
import type { SpreadsheetWorkbook } from '../spreadsheet/types';
import {
  adaptBiff8Sheet,
  adaptBiff8Workbook,
  attachBiff8Charts,
  attachBiff8DrawingImages,
} from './adapter';
import {
  createParseYieldState,
  yieldToBrowserIfNeeded,
} from './biff8/Biff8Reader';
import { parseBiff8Globals } from './biff8/parseGlobals';
import { parseBiff8Worksheet } from './biff8/parseWorksheet';
import { readBiff8ChartSubstream } from './chart/parseCharts';
import { XlsParseError } from './errors';
import type { Biff8Workbook } from './types';

/** 汇总XLS/BIFF8 解析当前步骤需要共享的上下文。 */
export type XlsCoreContext = {
  /** 在长任务检查点报告进度并响应取消信号。 */
  checkpoint(progress?: ParseProgress): Promise<void>;
  /** XlsCoreContext 处理完成后生成的输出结果。 */
  output?: XlsCoreOutput;
};

/** 描述 XlsCoreOutput 在 XLS/BIFF8 解析中的数据结构。 */
export type XlsCoreOutput = {
  /** 接收解析器产生的可移植资源分块。 */
  resource(resource: PortableResource): Promise<void>;
  /** XlsCoreOutput 当前关联的工作表。 */
  sheet(
    index: number,
    revision: number,
    sheet: SpreadsheetWorkbook['sheets'][number],
  ): Promise<void>;
};

/** 描述 XLS/BIFF8 解析产生的处理结果。 */
export type XlsCoreResult = {
  /** XlsCoreResult 当前关联的标准化工作簿。 */
  workbook: SpreadsheetWorkbook;
  /** XlsCoreResult 持有的图片、字体或对象 URL 等资源；文档释放时需同步清理。 */
  resources: PortableResource[];
};

/** 把输入映射为 `mapCfbError` 返回的结构。 */
function mapCfbError(error: CfbParseError) {
  const corruptedChain =
    error.code === 'CHAIN_CYCLE' ||
    error.code === 'CHAIN_TRUNCATED' ||
    error.code === 'SECTOR_OUT_OF_RANGE';
  return new XlsParseError(
    corruptedChain ? 'CORRUPTED_SECTOR_CHAIN' : 'INVALID_CFB',
    `XLS 容器解析失败：${error.message}`,
  );
}

/** 判断 `hasVbaStorage` 对应的条件是否成立。 */
function hasVbaStorage(
  entries: Awaited<ReturnType<typeof parseCfb>>['entries'],
) {
  // 宏只做目录级检测，绝不读取、反编译或执行 VBA 字节码。
  return entries.some((entry) => {
    const normalized = entry.path.toLowerCase();
    return (
      normalized.includes('/vba/') ||
      entry.name.toLowerCase() === '_vba_project_cur' ||
      entry.name.toLowerCase() === 'vba'
    );
  });
}

/** 判断 `hasSheetEnhancements` 对应的条件是否成立。 */
function hasSheetEnhancements(
  initial: SpreadsheetWorkbook['sheets'][number] | undefined,
  current: SpreadsheetWorkbook['sheets'][number],
) {
  return (
    current.images.length > 0 ||
    current.charts.length > 0 ||
    (Boolean(initial) &&
      (current.rowCount !== initial?.rowCount ||
        current.columnCount !== initial?.columnCount))
  );
}

/** 解析 XLS 二进制并返回环境无关的工作簿与资源。 */
export async function parseXlsCore(
  input: ArrayBuffer | Uint8Array,
  context: XlsCoreContext,
): Promise<XlsCoreResult> {
  const yieldState = createParseYieldState(8, () => context.checkpoint());
  await context.checkpoint({
    stage: 'container',
    percent: 0.05,
    message: '正在读取 XLS 复合文档',
  });
  let cfb: Awaited<ReturnType<typeof parseCfb>>;
  try {
    cfb = await parseCfb(input, {
      yieldIfNeeded: () => yieldToBrowserIfNeeded(yieldState),
    });
  } catch (error) {
    if (error instanceof CfbParseError) throw mapCfbError(error);
    throw error;
  }

  const workbookStream = cfb.getStream('Workbook', 'Book');
  if (!workbookStream) {
    throw new XlsParseError(
      'MISSING_WORKBOOK_STREAM',
      'XLS 文件缺少 Workbook 数据流',
    );
  }
  await context.checkpoint({
    stage: 'structure',
    percent: 0.18,
    message: '正在解析工作簿结构',
  });
  const globals = await parseBiff8Globals(workbookStream, yieldState);
  const warnings = [...globals.warnings];
  globals.hasVba = hasVbaStorage(cfb.entries);
  if (globals.hasVba) {
    warnings.push({
      code: 'VBA_DETECTED',
      message: '检测到 VBA 工程；预览器不会读取或执行宏代码',
    });
  }

  const descriptorsByOffset = [...globals.sheets].sort(
    (left, right) => left.streamOffset - right.streamOffset,
  );
  const workbook: Biff8Workbook = {
    globals,
    worksheets: [],
    chartSheets: [],
    warnings,
  };
  const initialSheets = new Map<
    string,
    SpreadsheetWorkbook['sheets'][number]
  >();
  const sheetIndexById = new Map(
    globals.sheets.map((descriptor, index) => [descriptor.id, index]),
  );
  for (let index = 0; index < descriptorsByOffset.length; index += 1) {
    const descriptor = descriptorsByOffset[index];
    const endOffset =
      descriptorsByOffset[index + 1]?.streamOffset ?? workbookStream.length;
    if (descriptor.type === 'worksheet') {
      const sheet = await parseBiff8Worksheet(
        workbookStream,
        descriptor,
        globals,
        endOffset,
        yieldState,
      );
      workbook.worksheets.push(sheet);
      warnings.push(...sheet.warnings);
    } else if (descriptor.type === 'chart') {
      workbook.chartSheets.push({
        descriptor,
        substream: readBiff8ChartSubstream(
          workbookStream,
          descriptor,
          endOffset,
        ),
      });
    }
    const initialSheet = adaptBiff8Sheet(workbook, descriptor);
    const sheetIndex = sheetIndexById.get(descriptor.id);
    if (context.output && initialSheet && sheetIndex !== undefined) {
      initialSheets.set(initialSheet.id, initialSheet);
      await context.output.sheet(sheetIndex, 0, initialSheet);
    }
    await context.checkpoint({
      stage: 'content',
      completed: index + 1,
      total: descriptorsByOffset.length,
      percent:
        0.25 + ((index + 1) / Math.max(1, descriptorsByOffset.length)) * 0.5,
      message: `正在解析工作表 ${index + 1}/${descriptorsByOffset.length}`,
    });
  }

  const target = adaptBiff8Workbook(workbook);
  const resources: PortableResource[] = [];
  await context.checkpoint({
    stage: 'resources',
    percent: 0.8,
    message: '正在处理图片和图表',
  });
  await attachBiff8DrawingImages(target, workbook, {
    add: async (resource) => {
      if (context.output) await context.output.resource(resource);
      else resources.push(resource);
      return createResourceReference(resource.id);
    },
  });
  attachBiff8Charts(target, workbook);
  if (context.output) {
    for (const sheet of target.sheets) {
      const initial = initialSheets.get(sheet.id);
      const sheetIndex = sheetIndexById.get(sheet.id);
      if (sheetIndex !== undefined && hasSheetEnhancements(initial, sheet)) {
        await context.output.sheet(sheetIndex, 1, sheet);
      }
    }
  }
  await context.checkpoint({
    stage: 'assembling',
    percent: 0.95,
    message: '正在组装工作簿',
  });
  return { workbook: target, resources };
}
