import { yieldToMainThread } from '../../formats/word-pages/scheduler';
import type { OfficeArchiveReader } from '../../shared/ooxml/OfficeArchiveReader';
import { readOfficeXmlEvents } from '../../shared/ooxml/OfficeXmlEventReader';
import {
  childByLocalName,
  descendantByLocalName,
  matchesLocalName,
} from '../../shared/ooxml/xml';
import { loadDocxPackageContext } from './DocxPackageContext';
import {
  docxBlockParseOperations,
  markDocxTitle,
  normalizeDocxPageContents,
  readDocxBodyPage,
} from './parseDocx';
import { parseDocxBlock, type DocxBlockParseResult } from './parseDocxBlock';
import type { DocxBlock, DocxImage, DocxPage, DocxPageContent } from './types';

type DocxSourceMetadata = {
  page: DocxPage;
  preserveSectionPagination: boolean;
};

export type DocxSourceOutput = {
  metadata(metadata: DocxSourceMetadata): void | Promise<void>;
  page(page: DocxPageContent): void | Promise<void>;
  complete(result: {
    title: string;
    images: DocxImage[];
  }): void | Promise<void>;
};

function createXmlDocument(rootName: string) {
  return new DOMParser().parseFromString(
    `<${rootName}></${rootName}>`,
    'application/xml',
  );
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
  const defaultPage = readDocxBodyPage(profile.bodyNode);
  const defaultSection = childByLocalName(profile.bodyNode, 'sectPr');
  const defaultRegions = defaultSection
    ? docxBlockParseOperations.readSectionRegions(defaultSection, context)
    : {};
  await output.metadata({
    page: defaultPage,
    preserveSectionPagination: profile.preserveSectionPagination,
  });

  let currentBlocks: DocxBlock[] = [];
  let previousBoundaryWasExplicit = false;
  let bodyIndex = 0;
  let pageIndex = 0;
  const titleCandidates: DocxBlock[] = [];
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
      pageIndex += 1;
      await output.page({
        ...normalized,
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
  });
}
