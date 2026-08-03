import { disposeDocDocument, type DocDocument } from './doc/types';
import type { DocxDocument } from './docx/types';
import type { OfficeParseOptions } from './parsing';
import { createOfficeParseSession } from './parsing';
import type { PptxDocument } from './pptx/types';
import { disposePresentationDocument } from './presentation/dispose';
import type { PresentationDocument } from './presentation/types';
import { disposeDocumentSession } from './session';
import {
  disposeSpreadsheetWorkbook,
  type SpreadsheetWorkbook,
} from './spreadsheet/types';

// 组件入口只关心“文件类型 + 解析结果”，具体格式的包结构解析都收敛在各自 service 中。
export {
  detectPreviewKind,
  getPreviewFamily,
  isPresentationPreviewKind,
  isSpreadsheetPreviewKind,
  isSupportedOfficeFileName,
  isWordPreviewKind,
  SUPPORTED_OFFICE_EXTENSIONS,
} from './parsing/detectPreviewKind';
export type { PreviewFamily, PreviewKind } from './parsing/detectPreviewKind';

/** 使用格式判别字段关联对应的标准化文档或工作簿结果。 */
export type ParsedOfficeFile =
  | {
      /** 当前模型对应的 Office 内容类型。 */
      kind: 'pptx';
      /** 当前处理的标准化文档模型。 */
      document: PptxDocument;
    }
  | {
      /** 当前模型对应的 Office 内容类型。 */
      kind: 'ppt';
      /** 当前处理的标准化文档模型。 */
      document: PresentationDocument;
    }
  | {
      /** 当前模型对应的 Office 内容类型。 */
      kind: 'xlsx';
      /** 当前处理的标准化工作簿。 */
      workbook: SpreadsheetWorkbook;
    }
  | {
      /** 当前模型对应的 Office 内容类型。 */
      kind: 'xls';
      /** 当前处理的标准化工作簿。 */
      workbook: SpreadsheetWorkbook;
    }
  | {
      /** 当前模型对应的 Office 内容类型。 */
      kind: 'docx';
      /** 当前处理的标准化文档模型。 */
      document: DocxDocument;
    }
  | {
      /** 当前模型对应的 Office 内容类型。 */
      kind: 'doc';
      /** 当前处理的标准化文档模型。 */
      document: DocDocument;
    };

/** 释放完整解析结果持有的格式资源和附加文档会话。 */
export async function disposeParsedOfficeFile(
  parsed: ParsedOfficeFile | undefined,
) {
  if (!parsed) return;

  const owner =
    parsed.kind === 'xls' || parsed.kind === 'xlsx'
      ? parsed.workbook
      : parsed.document;
  if (parsed.kind === 'xls' || parsed.kind === 'xlsx') {
    disposeSpreadsheetWorkbook(parsed.workbook);
  } else if (parsed.kind === 'ppt' || parsed.kind === 'pptx') {
    disposePresentationDocument(parsed.document);
  } else if (parsed.kind === 'doc') {
    disposeDocDocument(parsed.document);
  }
  await disposeDocumentSession(owner);
}

/** 识别 Office 文件格式并解析为标准预览模型。 */
export async function parseOfficeFile(
  file: File,
  options?: OfficeParseOptions,
): Promise<ParsedOfficeFile> {
  const session = createOfficeParseSession(file, options);
  try {
    return await session.result;
  } finally {
    session.dispose();
  }
}
