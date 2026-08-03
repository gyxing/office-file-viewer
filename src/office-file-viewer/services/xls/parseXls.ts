import { XlsDocumentAssembler } from '../parsing/assembly/DocumentAssembler';
import { ResourceRegistry } from '../parsing/assembly/ResourceRegistry';
import type { OfficeFormatParser } from '../parsing/formatParserRegistry';
import { throwIfParseAborted } from '../parsing/runtime/types';
import { createOfficeDocumentSession } from '../session';
import type { SpreadsheetWorkbook } from '../spreadsheet/types';
import { parseXlsCore } from './parseXlsCore';

/** 通过统一运行时合同解析 XLS，并逐张输出工作表和资源。 */
export const runXlsParser: OfficeFormatParser = async (
  file,
  { signal },
  sink,
) => {
  sink.progress({
    stage: 'reading',
    percent: 0.01,
    message: '正在读取 XLS 文件',
  });
  const input = await file.arrayBuffer();
  throwIfParseAborted(signal);
  const result = await parseXlsCore(input, {
    checkpoint: async (progress) => {
      throwIfParseAborted(signal);
      if (progress) sink.progress(progress);
    },
    output: {
      resource: async (resource) => {
        throwIfParseAborted(signal);
        await sink.resource(resource);
      },
      sheet: async (index, revision, sheet) => {
        throwIfParseAborted(signal);
        await sink.sheet(index, revision, sheet);
      },
    },
  });
  await sink.complete(result.workbook.warnings);
};

/** 在纯浏览器中解析未加密的 Excel 97–2003 BIFF8 工作簿。 */
export async function parseXls(file: File): Promise<SpreadsheetWorkbook> {
  const documentSession = createOfficeDocumentSession();
  const assembler = new XlsDocumentAssembler(new ResourceRegistry());
  let target: SpreadsheetWorkbook | undefined;
  try {
    await runXlsParser(
      file,
      {
        documentSessionId: documentSession.id,
        signal: documentSession.signal,
      },
      {
        progress: () => undefined,
        resource: (resource) => assembler.addResource(resource),
        sheet: async (index, revision, sheet) =>
          assembler.addSheet(index, revision, sheet),
        presentationMetadata: async () => {
          throw new Error('XLS 主线程运行时返回了错误的演示文稿元数据');
        },
        slide: async () => {
          throw new Error('XLS 主线程运行时返回了错误的幻灯片分块');
        },
        documentMetadata: async () => {
          throw new Error('XLS 主线程运行时返回了错误的文档元数据');
        },
        documentBlocks: async () => {
          throw new Error('XLS 主线程运行时返回了错误的正文分块');
        },
        parsed: async () => {
          throw new Error('XLS 主线程运行时返回了错误的完整文档消息');
        },
        complete: (warnings) => {
          assembler.setWarnings(warnings);
          target = assembler.complete();
        },
        error: () => undefined,
      },
    );
    if (!target) throw new Error('XLS 解析未返回完整工作簿');
    return target;
  } catch (error) {
    assembler.dispose();
    throw error;
  } finally {
    await documentSession.dispose();
  }
}
