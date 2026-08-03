import { PptDocumentAssembler } from '../parsing/assembly/DocumentAssembler';
import { ResourceRegistry } from '../parsing/assembly/ResourceRegistry';
import type { OfficeFormatParser } from '../parsing/formatParserRegistry';
import { throwIfParseAborted } from '../parsing/runtime/types';
import type { PresentationDocument } from '../presentation/types';
import { parsePptCore } from './parsePptCore';

function createYieldIfNeeded() {
  let deadline = Date.now() + 12;
  return async () => {
    if (Date.now() < deadline) return;
    await new Promise<void>((resolve) => {
      setTimeout(resolve, 0);
    });
    deadline = Date.now() + 12;
  };
}

/** 通过统一运行时合同解析 PPT，并按页输出幻灯片和资源。 */
export const runPptParser: OfficeFormatParser = async (
  file,
  { signal },
  sink,
) => {
  sink.progress({
    stage: 'reading',
    percent: 0.01,
    message: '正在读取 PPT 文件',
  });
  const input = await file.arrayBuffer();
  throwIfParseAborted(signal);
  await parsePptCore(input, {
    checkpoint: async (progress) => {
      throwIfParseAborted(signal);
      if (progress) sink.progress(progress);
    },
    output: {
      resource: async (resource) => {
        throwIfParseAborted(signal);
        await sink.resource(resource);
      },
      presentationMetadata: async (metadata) => {
        throwIfParseAborted(signal);
        await sink.presentationMetadata(metadata);
      },
      slide: async (index, slide) => {
        throwIfParseAborted(signal);
        await sink.slide(index, slide);
      },
    },
  });
  await sink.complete();
};

/** 在纯浏览器中解析未加密的 PowerPoint 97–2003 PPT 文件。 */
export async function parsePpt(file: File): Promise<PresentationDocument> {
  const assembler = new PptDocumentAssembler(new ResourceRegistry());
  const yieldIfNeeded = createYieldIfNeeded();
  try {
    const result = await parsePptCore(await file.arrayBuffer(), {
      checkpoint: yieldIfNeeded,
    });
    for (const resource of result.resources) {
      await assembler.addResource(resource);
    }
    const { width, height, theme, warnings, slides } = result.document;
    assembler.setMetadata({ width, height, theme, warnings });
    slides.forEach((slide, index) => assembler.addSlide(index, slide));
    return assembler.complete();
  } catch (error) {
    assembler.dispose();
    throw error;
  }
}
