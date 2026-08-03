import { DocDocumentAssembler } from '../parsing/assembly/DocumentAssembler';
import { ResourceRegistry } from '../parsing/assembly/ResourceRegistry';
import type { OfficeFormatParser } from '../parsing/formatParserRegistry';
import type { RuntimeSink } from '../parsing/runtime/types';
import { throwIfParseAborted } from '../parsing/runtime/types';
import { OFFICE_LARGE_FILE_THRESHOLDS } from '../performance/officePerformanceThresholds';
import { chunkDocBlocks, documentMetadataFromDoc } from './chunkDocBlocks';
import { parseDocCore, type DocCoreContext } from './parseDocCore';
import { parseDocRandomAccess } from './parseDocRandomAccess';
import type { DocDocument } from './types';

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

function createDocParserCheckpoint(signal: AbortSignal, sink: RuntimeSink) {
  let deadline = Date.now() + 12;
  return async (progress?: Parameters<RuntimeSink['progress']>[0]) => {
    throwIfParseAborted(signal);
    if (progress) sink.progress(progress);
    if (Date.now() < deadline) return;
    await new Promise<void>((resolve) => {
      setTimeout(resolve, 0);
    });
    throwIfParseAborted(signal);
    deadline = Date.now() + 12;
  };
}

/** 通过统一运行时合同解析 DOC/WPS，并按块输出正文和资源。 */
export const runDocParser: OfficeFormatParser = async (
  file,
  { signal },
  sink,
) => {
  sink.progress({
    stage: 'reading',
    percent: 0.01,
    message: '正在读取 DOC/WPS 文件',
  });
  const parseContext: DocCoreContext = {
    fileName: file.name,
    checkpoint: createDocParserCheckpoint(signal, sink),
    output: {
      resource: async (resource) => {
        throwIfParseAborted(signal);
        await sink.resource(resource);
      },
      documentMetadata: async (metadata) => {
        throwIfParseAborted(signal);
        await sink.documentMetadata(metadata);
      },
      documentBlocks: async (startIndex, blocks) => {
        throwIfParseAborted(signal);
        await sink.documentBlocks(startIndex, blocks);
      },
    },
  };

  if (file.size >= OFFICE_LARGE_FILE_THRESHOLDS.cfbFileBytes) {
    await parseDocRandomAccess(file, parseContext, signal);
  } else {
    const input = await file.arrayBuffer();
    throwIfParseAborted(signal);
    await parseDocCore(input, parseContext);
  }
  await sink.complete();
};

/** 保留原有 DOC/WPS 主线程解析入口，并复用跨运行时组装流程。 */
export async function parseDoc(file: File): Promise<DocDocument> {
  const assembler = new DocDocumentAssembler(new ResourceRegistry());
  const yieldIfNeeded = createYieldIfNeeded();
  try {
    const result = await parseDocCore(await file.arrayBuffer(), {
      fileName: file.name,
      checkpoint: yieldIfNeeded,
    });
    assembler.setMetadata(documentMetadataFromDoc(result.document));
    for (const chunk of chunkDocBlocks(result.document.blocks)) {
      assembler.addBlocks(chunk.startIndex, chunk.blocks);
    }
    for (const resource of result.resources) {
      await assembler.addResource(resource);
    }
    return assembler.complete();
  } catch (error) {
    assembler.dispose();
    throw error;
  }
}
