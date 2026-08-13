import type { PresentationDocument } from '../../presentation/types';
import type { ParsedOfficeFile } from '../../preview';
import { PptDocumentAssembler } from '../assembly/DocumentAssembler';
import { ResourceRegistry } from '../assembly/ResourceRegistry';
import {
  createMaterializedPreviewHandle,
  createMaterializedPreviewState,
  createOfficeFormatRuntimeSink,
  retainMaterializedPreview,
  type OfficeFormatSessionAdapterFactory,
} from './types';

/** 创建 PPT/PPTX 共用的增量组装与失败恢复适配器。 */
export const createPresentationSessionAdapter: OfficeFormatSessionAdapterFactory =
  (context) => {
    const assembler = new PptDocumentAssembler(new ResourceRegistry());
    let parsedResult: ParsedOfficeFile | undefined;

    const createParsedPresentation = (
      document: PresentationDocument,
    ): ParsedOfficeFile =>
      context.kind === 'pptx'
        ? { kind: 'pptx', document }
        : { kind: 'ppt', document };

    const emitSnapshot = () => {
      if (!context.enablePartial || !assembler.hasRenderableContent()) return;
      context.emitPartial(
        createMaterializedPreviewState(
          context,
          createParsedPresentation(assembler.snapshot()),
        ),
      );
    };

    const sink = createOfficeFormatRuntimeSink('演示文稿', context, {
      resource: async (resource) => assembler.addResource(resource),
      presentationMetadata: async (metadata) => {
        assembler.setMetadata(metadata);
        emitSnapshot();
      },
      slide: async (index, slide) => {
        assembler.addSlide(index, slide);
        emitSnapshot();
      },
      parsed: async (parsed) => {
        parsedResult = parsed;
      },
      complete: async () => {
        // 主线程入口已交付完整模型时不能再次消费 Worker 分块。
        if (parsedResult) return;
        parsedResult = createParsedPresentation(assembler.complete());
      },
    });

    return {
      sink,
      async finish() {
        if (!parsedResult) throw new Error('解析运行时未返回演示文稿结果');
        return createMaterializedPreviewHandle(context, parsedResult);
      },
      async recoverPartial() {
        if (!context.enablePartial || !assembler.hasRenderableContent()) {
          return undefined;
        }
        return retainMaterializedPreview(
          context,
          createParsedPresentation(assembler.completePartial()),
        );
      },
      dispose: () => assembler.dispose(),
    };
  };
