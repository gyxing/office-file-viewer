import type { DocxBlock, DocxPageContent } from '../../../docx/types';
import type { SlideModel } from '../../../presentation/types';
import type {
  SpreadsheetSheet,
  SpreadsheetWarning,
} from '../../../spreadsheet/types';
import type {
  PortableDocxMetadata,
  PortablePresentationMetadata,
  PortableSpreadsheetMetadata,
} from '../../protocol/messages';
import type { ParseProgress } from '../../types';
import { createParseAbortError } from '../types';

/** 单条 DOCX 正文消息包含的最大块数，避免大型文档形成巨型克隆任务。 */
const DOCX_BLOCK_CHUNK_SIZE = 32;
/** 单条 DOCX 页面消息包含的最大页数。 */
const DOCX_PAGE_CHUNK_SIZE = 8;

/** OOXML 物化解析在 Worker 内部使用的格式集合。 */
export type OoxmlMaterializedKind = 'docx' | 'xlsx' | 'pptx';

/** 接收 OOXML 物化解析产生的可克隆分块。 */
export type OoxmlMaterializedOutput = {
  /** 接收解析阶段进度。 */
  progress(progress: ParseProgress): void;
  /** 接收 DOCX 正文和页面之外的文档元数据。 */
  docxMetadata(metadata: PortableDocxMetadata): Promise<void>;
  /** 接收连续的 DOCX 正文块。 */
  docxBlocks(startIndex: number, blocks: DocxBlock[]): Promise<void>;
  /** 接收连续的 DOCX 页面。 */
  docxPages(startIndex: number, pages: DocxPageContent[]): Promise<void>;
  /** 接收工作表集合之外的工作簿元数据。 */
  spreadsheetMetadata(metadata: PortableSpreadsheetMetadata): Promise<void>;
  /** 接收单张工作表。 */
  sheet(index: number, sheet: SpreadsheetSheet): Promise<void>;
  /** 接收幻灯片集合之外的演示文稿元数据。 */
  presentationMetadata(metadata: PortablePresentationMetadata): Promise<void>;
  /** 接收单张幻灯片。 */
  slide(index: number, slide: SlideModel): Promise<void>;
};

function throwIfAborted(signal: AbortSignal) {
  if (signal.aborted) throw createParseAbortError();
}

async function parseDocxMaterialized(
  file: File,
  signal: AbortSignal,
  output: OoxmlMaterializedOutput,
) {
  const { parseDocx } = await import('../../../docx/parseDocx');
  const document = await parseDocx(file, signal);
  throwIfAborted(signal);
  const { blocks, pages, ...metadata } = document;
  await output.docxMetadata(metadata);
  for (let start = 0; start < blocks.length; start += DOCX_BLOCK_CHUNK_SIZE) {
    throwIfAborted(signal);
    await output.docxBlocks(
      start,
      blocks.slice(start, start + DOCX_BLOCK_CHUNK_SIZE),
    );
  }
  const pageList = pages ?? [];
  for (let start = 0; start < pageList.length; start += DOCX_PAGE_CHUNK_SIZE) {
    throwIfAborted(signal);
    await output.docxPages(
      start,
      pageList.slice(start, start + DOCX_PAGE_CHUNK_SIZE),
    );
  }
}

async function parseXlsxMaterialized(
  file: File,
  signal: AbortSignal,
  output: OoxmlMaterializedOutput,
): Promise<SpreadsheetWarning[] | undefined> {
  const { parseXlsx } = await import('../../../xlsx/parseXlsx');
  const workbook = await parseXlsx(file, signal);
  throwIfAborted(signal);
  const { sheets, warnings, resources: _resources, ...metadata } = workbook;
  await output.spreadsheetMetadata(metadata);
  for (let index = 0; index < sheets.length; index += 1) {
    throwIfAborted(signal);
    await output.sheet(index, sheets[index]);
  }
  return warnings;
}

async function parsePptxMaterialized(
  file: File,
  signal: AbortSignal,
  output: OoxmlMaterializedOutput,
) {
  const { parsePptx } = await import('../../../pptx/parsePptx');
  const document = await parsePptx(file, signal);
  throwIfAborted(signal);
  const { slides, resources: _resources, ...metadata } = document;
  await output.presentationMetadata(metadata);
  for (let index = 0; index < slides.length; index += 1) {
    throwIfAborted(signal);
    await output.slide(index, slides[index]);
  }
}

/** 在 Worker 中按格式加载 OOXML 解析核心并输出结构化克隆分块。 */
export async function runOoxmlMaterializedParser(
  file: File,
  kind: OoxmlMaterializedKind,
  signal: AbortSignal,
  output: OoxmlMaterializedOutput,
): Promise<SpreadsheetWarning[] | undefined> {
  output.progress({
    stage: 'content',
    percent: 0.05,
    message: '正在解析文件',
  });
  if (kind === 'docx') {
    await parseDocxMaterialized(file, signal, output);
    return undefined;
  }
  if (kind === 'xlsx') {
    return parseXlsxMaterialized(file, signal, output);
  }
  await parsePptxMaterialized(file, signal, output);
  return undefined;
}
