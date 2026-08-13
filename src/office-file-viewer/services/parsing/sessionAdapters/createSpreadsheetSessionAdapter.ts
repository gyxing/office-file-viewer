import type { ParsedOfficeFile } from '../../preview';
import type { SpreadsheetWorkbook } from '../../spreadsheet/types';
import { XlsDocumentAssembler } from '../assembly/DocumentAssembler';
import { ResourceRegistry } from '../assembly/ResourceRegistry';
import {
  createMaterializedPreviewHandle,
  createMaterializedPreviewState,
  createOfficeFormatRuntimeSink,
  retainMaterializedPreview,
  type OfficeFormatSessionAdapterFactory,
} from './types';

/** 创建 XLS/XLSX 共用的增量组装与失败恢复适配器。 */
export const createSpreadsheetSessionAdapter: OfficeFormatSessionAdapterFactory =
  (context) => {
    const assembler = new XlsDocumentAssembler(new ResourceRegistry());
    let parsedResult: ParsedOfficeFile | undefined;

    const createParsedWorkbook = (
      workbook: SpreadsheetWorkbook,
    ): ParsedOfficeFile =>
      context.kind === 'xlsx'
        ? { kind: 'xlsx', workbook }
        : { kind: 'xls', workbook };

    const emitSnapshot = () => {
      if (!context.enablePartial || !assembler.hasRenderableContent()) return;
      context.emitPartial(
        createMaterializedPreviewState(
          context,
          createParsedWorkbook(assembler.snapshot()),
        ),
      );
    };

    const sink = createOfficeFormatRuntimeSink('电子表格', context, {
      resource: async (resource) => assembler.addResource(resource),
      sheet: async (index, revision, sheet) => {
        assembler.addSheet(index, revision, sheet);
        emitSnapshot();
      },
      spreadsheetMetadata: async (metadata) => {
        assembler.setMetadata(metadata);
      },
      parsed: async (parsed) => {
        parsedResult = parsed;
      },
      complete: async (warnings) => {
        // 主线程入口已交付完整模型时不能再次消费 Worker 分块。
        if (parsedResult) return;
        assembler.setWarnings(warnings);
        parsedResult = createParsedWorkbook(assembler.complete());
      },
    });

    return {
      sink,
      async finish() {
        if (!parsedResult) throw new Error('解析运行时未返回电子表格结果');
        return createMaterializedPreviewHandle(context, parsedResult);
      },
      async recoverPartial() {
        if (!context.enablePartial || !assembler.hasRenderableContent()) {
          return undefined;
        }
        return retainMaterializedPreview(
          context,
          createParsedWorkbook(assembler.completePartial()),
        );
      },
      dispose: () => assembler.dispose(),
    };
  };
