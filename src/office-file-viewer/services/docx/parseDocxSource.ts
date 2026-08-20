import type { OfficeArchiveReader } from '../../shared/ooxml/OfficeArchiveReader';
import { readOfficeXmlEvents } from '../../shared/ooxml/OfficeXmlEventReader';
import {
  childByLocalName,
  descendantByLocalName,
  matchesLocalName,
  parseXml,
} from '../../shared/ooxml/xml';
import type { OfficeSourcePreviewFactory } from '../parsing/formatParserRegistry';
import { WorkerWordPageSource } from '../parsing/runtime/source/WorkerWordPageSource';
import { yieldToMainThread } from '../performance/mainThreadScheduler';
import { disposeDocumentSession } from '../session';
import { WordRevisionRecordCollector } from '../word/review/WordRevisionRecordCollector';
import type { WordBookmarkTarget } from '../word/types';
import { collectDocxPageRevisionRecords } from './collectDocxRevisionRecords';
import { profileDocxArchive } from './docxArchiveProfile';
import { attachDocxFootnotesToPage } from './docxNoteReferences';
import { loadDocxPackageContext } from './DocxPackageContext';
import { DocxWordPageSource } from './DocxWordPageSource';
import { parseDocxBlock, type DocxBlockParseResult } from './parseDocxBlock';
import { buildDocxReviewDocument } from './parseDocxComments';
import {
  docxBlockParseOperations,
  markDocxTitle,
  normalizeDocxPageContents,
  readDocxBlockChildren,
  readDocxBodyPage,
} from './parseDocxContent';
import { parseDocxEndnotes } from './parseDocxEndnotes';
import { parseDocxFootnotes } from './parseDocxFootnotes';
import type {
  DocxBlock,
  DocxCharacterSpacingControl,
  DocxImage,
  DocxPage,
  DocxPageContent,
} from './types';

/** DOCX 按需数据源预先读取的文档元数据。 */
type DocxSourceMetadata = {
  /** 当前关联的页面模型。 */
  page: DocxPage;
  /** 是否保留源文档由节属性定义的物理分页。 */
  preserveSectionPagination: boolean;
  /** Word 对东亚标点和假名采用的字符间距压缩方式。 */
  characterSpacingControl?: DocxCharacterSpacingControl;
};

/** DOCX 按需数据源初始化后的输出。 */
export type DocxSourceOutput = {
  /** 接收流式 DOCX 解析产生的文档元数据。 */
  metadata(metadata: DocxSourceMetadata): void | Promise<void>;
  /** 接收流式 DOCX 解析产生的单页模型。 */
  page(page: DocxPageContent): void | Promise<void>;
  /** 通知接收方增量输出已经结束。 */
  complete(result: {
    /** 面向用户展示的标题。 */
    title: string;
    /** 当前文档或页面包含的图片资源。 */
    images: DocxImage[];
    /** 按源名称索引的文档内部书签。 */
    bookmarks: Record<string, WordBookmarkTarget>;
    /** 当前文档中可恢复的批注、修订、脚注和尾注。 */
    review: ReturnType<typeof buildDocxReviewDocument>;
    /** 可按正文引用呈现的脚注和尾注正文。 */
    notes: import('./types').DocxNotes | undefined;
  }): void | Promise<void>;
};

function createXmlDocument(rootName: string) {
  return parseXml(`<${rootName}></${rootName}>`);
}

function appendAttributes(
  element: Element,
  attributes: ReadonlyMap<string, string>,
) {
  attributes.forEach((value, qualifiedName) => {
    if (qualifiedName === 'xmlns' || qualifiedName.startsWith('xmlns:')) {
      return;
    }
    const prefix = qualifiedName.includes(':')
      ? qualifiedName.split(':')[0]
      : undefined;
    if (prefix) {
      element.setAttributeNS(
        prefix === 'xml'
          ? 'http://www.w3.org/XML/1998/namespace'
          : `urn:office-prefix:${prefix}`,
        qualifiedName,
        value,
      );
    } else {
      element.setAttribute(qualifiedName, value);
    }
  });
}

/** 逐个生成 w:body 的直接子元素；任一时刻只保留一个局部 DOM。 */
async function* readBodyElements(
  reader: OfficeArchiveReader,
  signal?: AbortSignal,
): AsyncGenerator<Element> {
  const stream = await reader.openStream('word/document.xml', signal);
  let insideBody = false;
  let depth = 0;
  let fragment: XMLDocument | undefined;
  let stack: Element[] = [];

  for await (const event of readOfficeXmlEvents(stream, signal)) {
    if (!insideBody) {
      if (event.type === 'open' && event.localName === 'body') {
        insideBody = true;
      }
      continue;
    }
    if (event.type === 'open') {
      depth += 1;
      if (depth === 1) {
        fragment = createXmlDocument(event.localName);
        stack = [fragment.documentElement];
        appendAttributes(fragment.documentElement, event.attributes);
      } else {
        const element = fragment!.createElement(event.localName);
        appendAttributes(element, event.attributes);
        stack[stack.length - 1].appendChild(element);
        stack.push(element);
      }
      continue;
    }
    if (event.type === 'text') {
      if (stack.length && event.text) {
        stack[stack.length - 1].appendChild(
          fragment!.createTextNode(event.text),
        );
      }
      continue;
    }
    if (depth === 0 && event.localName === 'body') break;
    if (depth === 1) {
      if (fragment) yield fragment.documentElement;
      fragment = undefined;
      stack = [];
    } else {
      stack.pop();
    }
    depth -= 1;
  }
}

async function scanBodyProfile(
  reader: OfficeArchiveReader,
  signal?: AbortSignal,
) {
  const bodyDocument = createXmlDocument('body');
  const bodyNode = bodyDocument.documentElement;
  let preserveSectionPagination = false;
  for await (const child of readBodyElements(reader, signal)) {
    if (matchesLocalName(child, 'sectPr')) {
      bodyNode.appendChild(bodyDocument.importNode(child, true));
    } else if (
      matchesLocalName(child, 'p') &&
      descendantByLocalName(child, 'sectPr')
    ) {
      preserveSectionPagination = true;
    }
  }
  return { bodyNode, preserveSectionPagination };
}

/** 两次流式读取主 XML：首遍只取页面属性，次遍逐块解析并立即发布物理页。 */
export async function parseDocxSource(
  reader: OfficeArchiveReader,
  output: DocxSourceOutput,
  signal?: AbortSignal,
  resourceNamespace = 'docx',
) {
  const profile = await scanBodyProfile(reader, signal);
  const packageContext = await loadDocxPackageContext(
    reader,
    profile.bodyNode,
    resourceNamespace,
    signal,
  );
  const context = packageContext.parseContext;
  const readNoteBlocks = (
    container: Element,
    idPrefix: string,
    noteContext: typeof context,
  ) =>
    readDocxBlockChildren(container, idPrefix, noteContext, {
      insidePageRegion: true,
    });
  const entries = context.packageState.entries;
  const footnotes = parseDocxFootnotes(entries, context, readNoteBlocks);
  const endnotes = parseDocxEndnotes(entries, context, readNoteBlocks);
  const defaultPage = readDocxBodyPage(profile.bodyNode);
  const defaultSection = childByLocalName(profile.bodyNode, 'sectPr');
  const defaultRegions = defaultSection
    ? docxBlockParseOperations.readSectionRegions(defaultSection, context)
    : {};
  await output.metadata({
    page: defaultPage,
    preserveSectionPagination: profile.preserveSectionPagination,
    characterSpacingControl: context.characterSpacingControl,
  });

  let currentBlocks: DocxBlock[] = [];
  let previousBoundaryWasExplicit = false;
  let bodyIndex = 0;
  let pageIndex = 0;
  const titleCandidates: DocxBlock[] = [];
  const revisionCollector = new WordRevisionRecordCollector();
  const publishPage = async (page: DocxPage, regions = defaultRegions) => {
    if (!currentBlocks.length) return;
    const rawPage: DocxPageContent = {
      id: `docx-source-page-${pageIndex + 1}`,
      page,
      blocks: currentBlocks,
      ...regions,
    };
    currentBlocks = [];
    for (const normalized of normalizeDocxPageContents([rawPage])) {
      collectDocxPageRevisionRecords(
        [normalized],
        revisionCollector,
        pageIndex,
      );
      pageIndex += 1;
      await output.page({
        ...attachDocxFootnotesToPage(normalized, footnotes),
        id: `docx-page-${pageIndex}`,
      });
    }
  };

  for await (const child of readBodyElements(reader, signal)) {
    if (matchesLocalName(child, 'sectPr')) continue;
    const result: DocxBlockParseResult = parseDocxBlock({
      node: child,
      index: bodyIndex,
      context,
      defaultPage,
      previousBlock: currentBlocks[currentBlocks.length - 1],
      previousBoundaryWasExplicit,
      operations: docxBlockParseOperations,
    });
    previousBoundaryWasExplicit = result.previousBoundaryWasExplicit;
    for (const event of result.events) {
      if (event.type === 'blocks') {
        currentBlocks.push(...event.blocks);
        if (titleCandidates.length < 24) {
          titleCandidates.push(...event.blocks);
        }
      } else {
        await publishPage(event.page, event.regions ?? defaultRegions);
      }
    }
    bodyIndex += 1;
    if (bodyIndex % 32 === 0) await yieldToMainThread(signal);
  }
  await publishPage(defaultPage);
  await output.complete({
    title: markDocxTitle(titleCandidates),
    images: context.images,
    bookmarks: context.bookmarks,
    review: buildDocxReviewDocument(
      context.review,
      footnotes.length + endnotes.length,
      revisionCollector.toArray(),
    ),
    notes:
      footnotes.length || endnotes.length ? { footnotes, endnotes } : undefined,
  });
}

/** 仅在 DOCX 画像命中大文件阈值时创建流式分页预览源。 */
export const tryCreateDocxSourcePreview: OfficeSourcePreviewFactory = async (
  file,
  {
    documentSession,
    emitProgress,
    emitPartial,
    resourcePolicy,
    workerSourceClient,
  },
) => {
  emitProgress({
    stage: 'container',
    percent: 0.02,
    message: '正在读取 DOCX 包目录',
  });
  const archive = await profileDocxArchive(
    file,
    documentSession.signal,
    resourcePolicy,
  );
  if (archive.profile.mode !== 'lazy') {
    await archive.reader.close();
    return undefined;
  }

  if (workerSourceClient) {
    await archive.reader.close();
    const opened = await workerSourceClient.openSource(
      file,
      'docx',
      resourcePolicy,
      {
        signal: documentSession.signal,
        onProgress: emitProgress,
      },
    );
    if (!opened.available || opened.source.kind !== 'docx') return undefined;
    const source = new WorkerWordPageSource(workerSourceClient, opened.source);
    documentSession.register({ dispose: () => source.dispose() });
    documentSession.transferTo(source);
    const state = {
      sessionId: documentSession.id,
      previewKind: 'docx' as const,
      mode: 'source' as const,
      source,
      summary: source.getSummary(),
    };
    emitPartial(state);
    return {
      ...state,
      dispose: () => disposeDocumentSession(source),
    };
  }

  const source = new DocxWordPageSource({
    sessionId: documentSession.id,
    reader: archive.reader,
    signal: documentSession.signal,
  });
  documentSession.register({ dispose: () => source.dispose() });
  documentSession.transferTo(source);

  const createState = () => ({
    sessionId: documentSession.id,
    previewKind: 'docx' as const,
    mode: 'source' as const,
    source,
    summary: source.getSummary(),
  });
  const emitRenderablePartial = () => {
    if (source.hasRenderableContent()) emitPartial(createState());
  };

  try {
    await parseDocxSource(
      archive.reader,
      {
        metadata: (metadata) => {
          source.setMetadata(metadata);
          emitRenderablePartial();
        },
        page: async (page) => {
          await source.addSourcePage(page);
          emitRenderablePartial();
        },
        complete: (result) => {
          source.finishParsing(result);
          emitRenderablePartial();
        },
      },
      documentSession.signal,
      documentSession.id,
    );
    await source.waitForCompletion(documentSession.signal);
    if (!source.getSnapshot().pages.length) {
      throw new Error('DOCX PageSource 未生成可渲染页面');
    }
    const state = createState();
    return {
      ...state,
      dispose: () => disposeDocumentSession(source),
    };
  } catch (error) {
    if (source.hasRenderableContent()) {
      const summary = source.getSummary();
      try {
        source.finishParsing({
          title: summary.title,
          images: summary.images,
          bookmarks: summary.bookmarks ?? {},
          review: summary.review,
          notes: summary.notes,
        });
      } catch {
        // 已完成或已取消的 Source 保留当前可渲染快照即可。
      }
      emitRenderablePartial();
    }
    throw error;
  }
};
