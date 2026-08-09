import { readXml } from '../../shared/ooxml/archive';
import {
  childByLocalName,
  matchesLocalName,
  parseXml,
} from '../../shared/ooxml/xml';
import type { OfficeArchiveResourcePolicy } from '../../shared/resource/OfficeResourcePolicy';
import type { OfficeFormatParser } from '../parsing/formatParserRegistry';
import { throwIfParseAborted } from '../parsing/runtime/types';
import { loadDocxEntries } from './archive';
import { createDocxParseContext } from './docxParsingContext';
import { parseDocxBlock, type DocxBlockParseResult } from './parseDocxBlock';
import {
  applyDocxCoverTitleSpacing,
  docxBlockParseOperations,
  markDocxTitle,
  normalizeDocxPageContents,
  readDocxBodyPage,
  readDocxSectionPageRegions,
} from './parseDocxContent';
import type {
  DocxBlock,
  DocxDocument,
  DocxPage,
  DocxPageContent,
} from './types';

function throwIfDocxParseAborted(signal?: AbortSignal) {
  if (!signal?.aborted) return;
  const error = new Error('DOCX 解析已取消');
  error.name = 'AbortError';
  throw error;
}

/** 在正文分批边界让出主线程，使取消事件能够及时生效。 */
async function docxParseCheckpoint(signal?: AbortSignal) {
  throwIfDocxParseAborted(signal);
  await new Promise<void>((resolve) => {
    setTimeout(resolve, 0);
  });
  throwIfDocxParseAborted(signal);
}

/** 解析 DOCX 包并返回标准文档模型。 */
export async function parseDocx(
  file: File,
  signal?: AbortSignal,
  resourcePolicy?: OfficeArchiveResourcePolicy,
): Promise<DocxDocument> {
  // 解析顺序：包资源 -> 主题/样式 -> body 子节点，段落/表格内部再递归解析图片、图表和形状。
  throwIfDocxParseAborted(signal);
  const entries = await loadDocxEntries(file, { signal, resourcePolicy });
  await docxParseCheckpoint(signal);
  const documentXml = readXml(entries, 'word/document.xml');
  const documentDoc = parseXml(documentXml);
  const bodyNode = childByLocalName(documentDoc.documentElement, 'body');
  const preserveSectionPagination = Array.from(bodyNode?.children ?? []).some(
    (child) =>
      matchesLocalName(child, 'p') &&
      Boolean(childByLocalName(childByLocalName(child, 'pPr'), 'sectPr')),
  );
  const context = createDocxParseContext(entries, { bodyNode });

  const blocks: DocxBlock[] = [];
  const pages: DocxPageContent[] = [];
  let currentPageBlocks: DocxBlock[] = [];
  let previousBoundaryWasExplicit = false;
  const defaultSectPr = childByLocalName(bodyNode, 'sectPr');
  const defaultRegions = readDocxSectionPageRegions(defaultSectPr, context);
  const pushCurrentPage = (
    page: DocxPage,
    regions: ReturnType<typeof readDocxSectionPageRegions> = defaultRegions,
  ) => {
    if (!currentPageBlocks.length) return;
    pages.push({
      id: `docx-page-${pages.length + 1}`,
      page,
      blocks: currentPageBlocks,
      ...regions,
    });
    currentPageBlocks = [];
  };

  const bodyChildren = Array.from(bodyNode?.children ?? []);
  for (let index = 0; index < bodyChildren.length; index += 1) {
    if (index % 32 === 0) await docxParseCheckpoint(signal);
    const child = bodyChildren[index];
    const page = readDocxBodyPage(bodyNode);
    const result: DocxBlockParseResult = parseDocxBlock({
      node: child,
      index,
      context,
      defaultPage: page,
      previousBlock: currentPageBlocks[currentPageBlocks.length - 1],
      previousBoundaryWasExplicit,
      operations: docxBlockParseOperations,
    });
    previousBoundaryWasExplicit = result.previousBoundaryWasExplicit;
    for (const event of result.events) {
      if (event.type === 'blocks') {
        blocks.push(...event.blocks);
        currentPageBlocks.push(...event.blocks);
      } else {
        pushCurrentPage(event.page, event.regions ?? defaultRegions);
      }
    }
  }

  if (currentPageBlocks.length) {
    pushCurrentPage(readDocxBodyPage(bodyNode));
  }

  applyDocxCoverTitleSpacing(blocks);
  const normalizedPages = normalizeDocxPageContents(pages);
  const outline = blocks.flatMap((block) =>
    block.type === 'paragraph' && block.outlineLevel !== undefined && block.text
      ? [
          {
            id: `outline-${block.id}`,
            text: block.text,
            level: block.outlineLevel,
            targetBlockId: block.id,
          },
        ]
      : [],
  );

  throwIfDocxParseAborted(signal);
  return {
    title: markDocxTitle(blocks),
    page: normalizedPages[0]?.page ?? readDocxBodyPage(bodyNode),
    pages: normalizedPages,
    blocks,
    images: context.images,
    outline,
    bookmarks: context.bookmarks,
    preserveSectionPagination,
    characterSpacingControl: context.characterSpacingControl,
  };
}

/** 通过统一运行时合同解析 DOCX，并输出完整文档模型。 */
export const runDocxParser: OfficeFormatParser = async (
  file,
  { signal, resourcePolicy },
  sink,
) => {
  sink.progress({
    stage: 'content',
    percent: 0.05,
    message: '正在解析文件',
  });
  const document = await parseDocx(file, signal, resourcePolicy);
  throwIfParseAborted(signal);
  await sink.parsed({ kind: 'docx', document });
  await sink.complete();
};
