import { DocWordPageSource } from '../../doc/DocWordPageSource';
import { OFFICE_LARGE_FILE_THRESHOLDS } from '../../performance/officePerformanceThresholds';
import type { ParsedOfficeFile } from '../../preview';
import { disposeDocumentSession } from '../../session';
import { DocDocumentAssembler } from '../assembly/DocumentAssembler';
import { DocxDocumentAssembler } from '../assembly/DocxDocumentAssembler';
import { ResourceRegistry } from '../assembly/ResourceRegistry';
import type { OfficeFileViewerPreviewState } from '../internalTypes';
import {
  createMaterializedPreviewHandle,
  createMaterializedPreviewState,
  createOfficeFormatRuntimeSink,
  retainMaterializedPreview,
  type OfficeFormatSessionAdapterFactory,
} from './types';

/** 创建 DOC/DOCX 共用的增量组装、分页源与失败恢复适配器。 */
export const createWordSessionAdapter: OfficeFormatSessionAdapterFactory = (
  context,
) => {
  const useDocPageSource =
    context.enablePartial &&
    context.kind === 'doc' &&
    context.file.size >= OFFICE_LARGE_FILE_THRESHOLDS.cfbFileBytes;
  const documentAssembler =
    context.kind === 'doc' && !useDocPageSource
      ? new DocDocumentAssembler(new ResourceRegistry())
      : undefined;
  const docxAssembler =
    context.kind === 'docx' ? new DocxDocumentAssembler() : undefined;
  const docPageSource = useDocPageSource
    ? new DocWordPageSource({
        sessionId: context.documentSession.id,
        signal: context.documentSession.signal,
      })
    : undefined;
  let parsedResult: ParsedOfficeFile | undefined;

  const createDocSourceState = (): OfficeFileViewerPreviewState => {
    if (!docPageSource) throw new Error('DOC PageSource 尚未创建');
    return {
      sessionId: context.documentSession.id,
      previewKind: 'doc',
      mode: 'source',
      source: docPageSource,
      summary: docPageSource.getSummary(),
    };
  };

  const emitDocSourcePartial = () => {
    if (!context.enablePartial || !docPageSource?.hasRenderableContent()) {
      return;
    }
    context.emitPartial(createDocSourceState());
  };

  const emitDocPartial = () => {
    if (!context.enablePartial || !documentAssembler?.hasRenderableContent()) {
      return;
    }
    context.emitPartial(
      createMaterializedPreviewState(context, {
        kind: 'doc',
        document: documentAssembler.snapshot(),
      }),
    );
  };

  const emitDocxPartial = () => {
    if (!context.enablePartial || !docxAssembler?.hasRenderableContent()) {
      return;
    }
    context.emitPartial(
      createMaterializedPreviewState(context, {
        kind: 'docx',
        document: docxAssembler.snapshot(),
      }),
    );
  };

  const sink = createOfficeFormatRuntimeSink('文字文档', context, {
    resource: async (resource) => {
      const target = documentAssembler ?? docPageSource;
      if (!target) throw new Error('当前文字文档会话收到了资源分块');
      await target.addResource(resource);
    },
    documentMetadata: async (metadata) => {
      if (docPageSource) {
        docPageSource.setMetadata(metadata);
        emitDocSourcePartial();
        return;
      }
      if (!documentAssembler) {
        throw new Error('非 DOC 会话收到了文档元数据');
      }
      documentAssembler.setMetadata(metadata);
      emitDocPartial();
    },
    documentBlocks: async (startIndex, blocks) => {
      if (docPageSource) {
        await docPageSource.addBlocks(startIndex, blocks);
        emitDocSourcePartial();
        return;
      }
      if (!documentAssembler) {
        throw new Error('非 DOC 会话收到了正文分块');
      }
      documentAssembler.addBlocks(startIndex, blocks);
      emitDocPartial();
    },
    docxMetadata: async (metadata) => {
      if (!docxAssembler) {
        throw new Error('非 DOCX 会话收到了文档元数据');
      }
      docxAssembler.setMetadata(metadata);
    },
    docxBlocks: async (startIndex, blocks) => {
      if (!docxAssembler) {
        throw new Error('非 DOCX 会话收到了正文分块');
      }
      docxAssembler.addBlocks(startIndex, blocks);
      emitDocxPartial();
    },
    docxPages: async (startIndex, pages) => {
      if (!docxAssembler) {
        throw new Error('非 DOCX 会话收到了页面分块');
      }
      docxAssembler.addPages(startIndex, pages);
      emitDocxPartial();
    },
    parsed: async (parsed) => {
      parsedResult = parsed;
    },
    complete: async () => {
      // 主线程入口已交付完整模型时不能再次消费 Worker 分块。
      if (parsedResult) return;
      if (documentAssembler) {
        parsedResult = {
          kind: 'doc',
          document: documentAssembler.complete(),
        };
        return;
      }
      if (docxAssembler) {
        parsedResult = {
          kind: 'docx',
          document: docxAssembler.complete(),
        };
        return;
      }
      if (docPageSource) await docPageSource.complete();
    },
  });

  return {
    sink,
    async finish() {
      if (docPageSource) {
        if (!docPageSource.hasRenderableContent()) {
          throw new Error('DOC PageSource 未生成可渲染页面');
        }
        context.documentSession.transferTo(docPageSource);
        return {
          ...createDocSourceState(),
          dispose: () => disposeDocumentSession(docPageSource),
        };
      }
      if (!parsedResult) throw new Error('解析运行时未返回文字文档结果');
      return createMaterializedPreviewHandle(context, parsedResult);
    },
    async recoverPartial() {
      if (!context.enablePartial) return undefined;
      if (docxAssembler?.hasRenderableContent()) {
        return retainMaterializedPreview(context, {
          kind: 'docx',
          document: docxAssembler.completePartial(),
        });
      }
      if (documentAssembler?.hasRenderableContent()) {
        return retainMaterializedPreview(context, {
          kind: 'doc',
          document: documentAssembler.completePartial(),
        });
      }
      if (docPageSource?.hasRenderableContent()) {
        await docPageSource.complete();
        context.documentSession.transferTo(docPageSource);
        return createDocSourceState();
      }
      return undefined;
    },
    async dispose() {
      documentAssembler?.dispose();
      docxAssembler?.dispose();
      await docPageSource?.dispose();
    },
  };
};
