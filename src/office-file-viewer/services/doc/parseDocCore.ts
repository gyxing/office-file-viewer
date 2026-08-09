import { CFB_SIGNATURE, parseCfb } from '../../shared/binary/cfb';
import type {
  PortableDocMetadata,
  PortableResource,
} from '../parsing/protocol/messages';
import type { ParseProgress } from '../parsing/types';
import {
  buildDocBlocksFromSegments,
  buildDocBlocksFromText,
} from './buildDocBlocks';
import { buildDocDrawingParagraphAnchors } from './buildDocDrawingParagraphAnchors';
import {
  documentMetadataFromDoc,
  paragraphsFromDocBlocks,
} from './chunkDocBlocks';
import type { DocBlockBuildOptions } from './docParseTypes';
import { extractDocImages } from './extractDocImages';
import {
  extractDocDrawingCanvases,
  type DocDrawingTextBox,
} from './parseDocDrawingCanvas';
import {
  DEFAULT_DOC_PAGE,
  readDocBinaryContent,
  readDocFib,
  readDocStorySegments,
} from './readDocBinaryContent';
import {
  attachDocBookmarkMarkers,
  buildDocBookmarkTargets,
} from './readDocBookmarks';
import type {
  DocBlock,
  DocDocument,
  DocImage,
  DocParagraphBlock,
} from './types';

/** 汇总 DOC 二进制解析各步骤共享的上下文。 */
export type DocCoreContext = {
  /** 正在解析的原始文件名，用于格式识别和错误提示。 */
  fileName: string;
  /** 在长任务检查点报告进度并响应取消信号。 */
  checkpoint(progress?: ParseProgress): Promise<void>;
  /** 处理完成后生成的输出结果。 */
  output?: DocCoreOutput;
};

/** DOC 核心解析生成的文档及性能档案。 */
export type DocCoreOutput = {
  /** 接收解析器产生的可移植资源分块。 */
  resource(resource: PortableResource): Promise<void>;
  /** 接收文字文档的主体元数据。 */
  documentMetadata(metadata: PortableDocMetadata): Promise<void>;
  /** 接收文字文档的连续内容块。 */
  documentBlocks(startIndex: number, blocks: DocBlock[]): Promise<void>;
};

/** DOC 核心解析成功或失败的联合结果。 */
export type DocCoreResult = {
  /** 当前处理的标准化文档模型。 */
  document: DocDocument;
  /** 持有的图片、字体或对象 URL 等资源；文档释放时需同步清理。 */
  resources: PortableResource[];
};

/** 提供已经由随机 CFB Reader 读取的 DOC 核心流。 */
export type DocCoreStreamsInput = {
  /** WordDocument 主流。 */
  wordDocument: Uint8Array;
  /** FIB 指定的 0Table 或 1Table 流。 */
  tableStream: Uint8Array;
  /** 仅用于提取图片资源的相关流，通常为 WordDocument、Table 和 Data。 */
  imageStreams: Iterable<readonly [string, Uint8Array]>;
};

function isOleDoc(bytes: Uint8Array) {
  return CFB_SIGNATURE.every((value, index) => bytes[index] === value);
}

/** 判断输入是否为随机 CFB Reader 提供的 DOC 核心流。 */
function isDocCoreStreamsInput(
  input: ArrayBuffer | Uint8Array | DocCoreStreamsInput,
): input is DocCoreStreamsInput {
  return (
    typeof input === 'object' &&
    input !== null &&
    'wordDocument' in input &&
    'tableStream' in input
  );
}

async function parsePlainLikeDoc(
  bytes: Uint8Array,
  fileName: string,
  warnings: string[],
  options: DocBlockBuildOptions,
  output?: DocCoreOutput,
): Promise<DocDocument> {
  const fullText = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
  const isRtf = fullText.trimStart().startsWith('{\\rtf');
  const text = isRtf
    ? fullText
        .replace(/\\'[0-9a-f]{2}/gi, '')
        .replace(/\\[a-z]+-?\d* ?/gi, '')
        .replace(/[{}]/g, '')
    : fullText.replace(/<[^>]+>/g, ' ');

  warnings.push(
    isRtf
      ? '\u68c0\u6d4b\u5230 RTF \u5185\u5bb9\uff0c\u5df2\u6309\u7eaf\u6587\u672c\u964d\u7ea7\u9884\u89c8\u3002'
      : '\u68c0\u6d4b\u5230\u975e OLE DOC \u5185\u5bb9\uff0c\u5df2\u6309\u7eaf\u6587\u672c\u964d\u7ea7\u9884\u89c8\u3002',
  );
  const metadataDocument = buildDocDocument(fileName, [], [...warnings]);
  await output?.documentMetadata(documentMetadataFromDoc(metadataDocument));
  const blocks = await buildDocBlocksFromText(text, options);
  return buildDocDocument(fileName, blocks, warnings);
}

function buildDocDocument(
  fileName: string,
  blocks: DocBlock[],
  warnings: string[],
): DocDocument {
  const paragraphs = paragraphsFromDocBlocks(blocks);
  const title =
    paragraphs.find((paragraph) => paragraph.text)?.text ??
    (fileName || 'DOC \u6587\u6863');
  const images = [] as DocImage[];
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

  return {
    title,
    page: DEFAULT_DOC_PAGE,
    blocks,
    paragraphs,
    images,
    outline,
    bookmarks: buildDocBookmarkTargets(blocks),
    warnings,
  };
}

/** 消费已提取图片，确保引用它们的 DOC 元数据和正文不会先到达。 */
async function flushDocResources(
  resources: PortableResource[],
  output: DocCoreOutput | undefined,
) {
  if (!output || !resources.length) return;
  const pending = resources.splice(0);
  for (const resource of pending) {
    await output.resource(resource);
  }
}

/** 解析 DOC/WPS 二进制，并返回环境无关的文档与图片资源。 */
export async function parseDocCore(
  input: ArrayBuffer | Uint8Array | DocCoreStreamsInput,
  context: DocCoreContext,
): Promise<DocCoreResult> {
  // 非 OLE 文件按纯文本降级处理；OLE DOC 则解析 CFB、FIB、piece table 和样式 run。
  const streamsInput = isDocCoreStreamsInput(input) ? input : undefined;
  const bytes = streamsInput
    ? undefined
    : input instanceof Uint8Array
    ? input
    : new Uint8Array(input as ArrayBuffer);
  const warnings: string[] = [];
  const resources: PortableResource[] = [];

  if (bytes && !isOleDoc(bytes)) {
    await context.checkpoint({
      stage: 'content',
      percent: 0.8,
      message: '正在解析 DOC 纯文本内容',
    });
    const document = await parsePlainLikeDoc(
      bytes,
      context.fileName,
      warnings,
      {
        checkpoint: context.checkpoint,
        onBatch: context.output
          ? (startIndex, blocks) =>
              context.output!.documentBlocks(startIndex, blocks)
          : undefined,
      },
      context.output,
    );
    await context.output?.documentMetadata(documentMetadataFromDoc(document));
    return { document, resources };
  }

  await context.checkpoint({
    stage: 'container',
    percent: 0.05,
    message: '正在读取 DOC 复合文档',
  });
  const cfb = bytes
    ? await parseCfb(bytes, {
        yieldIfNeeded: () => context.checkpoint(),
        allowPartialFinalSector: true,
      })
    : undefined;
  const inputStreams = streamsInput
    ? [...streamsInput.imageStreams]
    : undefined;
  const dataStream =
    inputStreams?.find(
      ([streamName]) =>
        streamName.replace(/^.*[\\/]/, '').toLowerCase() === 'data',
    )?.[1] ?? cfb?.getStream('Data');
  const wordDocument =
    streamsInput?.wordDocument ?? cfb?.getStream('WordDocument');

  if (!wordDocument) {
    throw new Error(
      'DOC \u6587\u4ef6\u7f3a\u5c11 WordDocument \u6570\u636e\u6d41',
    );
  }

  const fib = readDocFib(wordDocument);
  const tableStream =
    streamsInput?.tableStream ?? cfb?.getStream(fib.tableStreamName);

  if (!tableStream) {
    throw new Error(
      `DOC \u6587\u4ef6\u7f3a\u5c11 ${fib.tableStreamName} \u6570\u636e\u6d41`,
    );
  }

  await context.checkpoint({
    stage: 'structure',
    percent: 0.25,
    message: '正在解析 DOC 文档结构',
  });
  const binaryContent = readDocBinaryContent({
    wordDocument,
    tableStream,
    dataStream,
    fib,
  });
  const {
    sections,
    normalStyle,
    outlineCatalog,
    numbering,
    paragraphRuns,
    drawingTextBoxRanges,
    bookmarks,
    bookmarkWarnings,
  } = binaryContent;
  const dominantSection = [...sections].sort(
    (left, right) =>
      right.charEnd - right.charStart - (left.charEnd - left.charStart),
  )[0];
  const documentPage = dominantSection?.page ?? DEFAULT_DOC_PAGE;
  const documentGridLinePitch = sections.find(
    (section) => section.gridLinePitch !== undefined,
  )?.gridLinePitch;
  const paragraphLineMultiplierCounts = paragraphRuns.reduce((counts, run) => {
    const multiplier = run.style?.lineHeightMultiplier;
    if (multiplier !== undefined && !run.inTable && !run.isTableOfContents) {
      counts.set(multiplier, (counts.get(multiplier) ?? 0) + 1);
    }
    return counts;
  }, new Map<number, number>());
  const dominantParagraphLineMultiplier = [
    ...paragraphLineMultiplierCounts.entries(),
  ].sort((left, right) => right[1] - left[1])[0]?.[0];
  const defaultLineMultiplier =
    normalStyle?.lineHeightMultiplier ??
    (normalStyle?.lineHeight !== undefined && normalStyle.lineHeight <= 4
      ? normalStyle.lineHeight
      : dominantParagraphLineMultiplier);
  const defaultGridLineHeight =
    documentGridLinePitch !== undefined && defaultLineMultiplier !== undefined
      ? documentGridLinePitch * defaultLineMultiplier
      : undefined;
  warnings.push(...outlineCatalog.warnings);
  warnings.push(...bookmarkWarnings);
  await context.checkpoint({
    stage: 'resources',
    percent: 0.5,
    message: '正在解析 DOC 图片资源',
  });
  const images = extractDocImages(inputStreams ?? cfb!.streams, resources);
  const headerStart = fib.ccpText + fib.ccpFtn;
  const headerText = readDocStorySegments(
    wordDocument,
    binaryContent,
    headerStart,
    headerStart + fib.ccpHdr,
  )
    .map((segment) => segment.text)
    .join('');
  // 图片流按 Word story 顺序排列；页眉存在图片锚点时，第一张图不应再分配给正文。
  const headerImage = headerText.includes('\u0001') ? images[0] : undefined;
  const bodyImages = headerImage ? images.slice(1) : images;
  const footerPageNumbers = /\u0013PAGE\b/.test(headerText);
  const textBoxStart =
    fib.ccpText +
    fib.ccpFtn +
    fib.ccpHdr +
    fib.ccpMcr +
    fib.ccpAtn +
    fib.ccpEdn;
  let textBoxes: DocDrawingTextBox[];
  if (drawingTextBoxRanges.length) {
    textBoxes = [];
    for (const range of drawingTextBoxRanges) {
      if (!range) {
        // PlcftxbxTxt 的空记录仍占用 ClientTextbox 索引，必须保留槽位。
        textBoxes.push({ text: '' });
        continue;
      }
      const scopedSegments = readDocStorySegments(
        wordDocument,
        binaryContent,
        textBoxStart + range.charStart,
        textBoxStart + range.charEnd,
      );
      const scopedBlocks = await buildDocBlocksFromSegments(
        scopedSegments,
        [],
        {
          checkpoint: context.checkpoint,
        },
      );
      const paragraphs = scopedBlocks.filter(
        (block): block is DocParagraphBlock => block.type === 'paragraph',
      );
      const styledParagraph =
        paragraphs.find((paragraph) => paragraph.text.trim()) ?? paragraphs[0];
      textBoxes.push({
        text: paragraphs
          .map((paragraph) => paragraph.text.trim())
          .filter(Boolean)
          .join('\n'),
        style: styledParagraph?.style,
      });
    }
  } else {
    const textBoxSegments = readDocStorySegments(
      wordDocument,
      binaryContent,
      textBoxStart,
      textBoxStart + fib.ccpTxbx,
    );
    const textBoxBlocks = await buildDocBlocksFromSegments(
      textBoxSegments,
      [],
      {
        checkpoint: context.checkpoint,
      },
    );
    textBoxes = textBoxBlocks.filter(
      (block): block is DocParagraphBlock => block.type === 'paragraph',
    );
  }
  const segments = attachDocBookmarkMarkers(
    readDocStorySegments(wordDocument, binaryContent, 0, fib.ccpText),
    bookmarks,
  );
  const drawingCanvases = extractDocDrawingCanvases(
    tableStream,
    fib,
    textBoxes,
    resources,
    {
      sections,
      displayPage: documentPage,
      paragraphAnchors: buildDocDrawingParagraphAnchors(segments),
    },
  );
  warnings.push(...drawingCanvases.warnings);
  const drawingImages = drawingCanvases.images;
  await flushDocResources(resources, context.output);
  const metadataDocument = buildDocDocument(
    context.fileName,
    [],
    [...warnings],
  );
  metadataDocument.page = documentPage;
  metadataDocument.images = [...drawingImages, ...images];
  metadataDocument.headerImage = headerImage;
  metadataDocument.footerPageNumbers = footerPageNumbers;
  await context.output?.documentMetadata(
    documentMetadataFromDoc(metadataDocument),
  );
  await context.checkpoint({
    stage: 'content',
    percent: 0.7,
    message: '正在解析 DOC 正文内容',
  });
  const hasStructuralTableRows = paragraphRuns.some((run) => run.tableRowEnd);
  // 只有文档实际提供行结束标志时才采用 PAPX 表格结构；旧 WPS/DOC 常只有 inTable，
  // 此时继续使用单元格分隔符回退，避免把表格及其后的正文吞成同一行。
  const contentSegments = hasStructuralTableRows
    ? segments
    : segments.map((segment) => ({
        ...segment,
        inTable: undefined,
        tableRowEnd: undefined,
      }));
  const blocks = await buildDocBlocksFromSegments(
    contentSegments,
    bodyImages,
    {
      checkpoint: context.checkpoint,
      numbering,
      defaultGridLineHeight,
      documentGridLineHeight: documentGridLinePitch,
      pageContentHeight:
        documentPage.minHeight -
        documentPage.marginTop -
        documentPage.marginBottom,
      onBatch: context.output
        ? (startIndex, batch) =>
            context.output!.documentBlocks(startIndex, batch)
        : undefined,
    },
    drawingCanvases.slots,
  );

  if (!blocks.length) {
    throw new Error(
      '\u8be5 DOC \u6587\u4ef6\u672a\u89e3\u6790\u5230\u53ef\u9884\u89c8\u6b63\u6587',
    );
  }

  warnings.push(
    drawingImages.length
      ? '已恢复 DOC/WPS 主文档中的 OfficeArt 绘图画布；分页仍由前端按源页面尺寸估算。'
      : images.length
      ? '当前为纯前端 DOC/WPS 降级预览，已提取到文档内图片，并按前端估算分页；暂未恢复精确锚点和复杂样式。'
      : '当前为纯前端 DOC/WPS 降级预览，已按前端估算分页；暂不还原复杂样式和图片锚点。',
  );
  await context.checkpoint({
    stage: 'assembling',
    percent: 0.95,
    message: '正在组装 DOC 文档',
  });
  const document = buildDocDocument(context.fileName, blocks, warnings);
  document.page = documentPage;
  document.images = [...drawingImages, ...images];
  document.headerImage = headerImage;
  document.footerPageNumbers = footerPageNumbers;
  await context.output?.documentMetadata(documentMetadataFromDoc(document));
  return { document, resources };
}
