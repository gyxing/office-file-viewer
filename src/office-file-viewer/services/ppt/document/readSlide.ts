import type { ThemeModel } from '../../presentation/types';
import { PPT_RECORD } from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import { parsePptDrawing } from '../drawing';
import { parsePptComments } from '../parsePptComments';
import { parsePptTransition } from '../parsePptTransitions';
import type {
  PptEditChain,
  PptParseContext,
  PptSlideDescriptor,
  PptSlideModel,
} from '../types';

/** 读取一页幻灯片的母版引用、文本与 OfficeArt 绘图。 */
export async function readPptSlide(
  documentStream: Uint8Array,
  editChain: PptEditChain,
  descriptor: PptSlideDescriptor,
  width: number,
  height: number,
  theme: ThemeModel,
  fonts: Map<number, string>,
  context: PptParseContext,
): Promise<PptSlideModel | undefined> {
  const offset = editChain.persistOffsets.get(descriptor.persistId);
  if (offset === undefined) {
    context.warnings.push({
      code: 'PPT_SLIDE_MISSING',
      message: `持久化目录中缺少第 ${descriptor.index} 页`,
      slideIndex: descriptor.index,
    });
    return undefined;
  }

  const record = new PptRecordReader(
    documentStream,
    offset,
    documentStream.length,
  ).readRecord();
  if (!record || record.type !== PPT_RECORD.SLIDE) {
    context.warnings.push({
      code: 'PPT_SLIDE_CORRUPT',
      message: `第 ${descriptor.index} 页不是有效的 SlideContainer`,
      slideIndex: descriptor.index,
      offset,
    });
    return undefined;
  }

  let masterId: number | undefined;
  let drawing: Uint8Array | undefined;
  const children = new PptRecordReader(
    documentStream,
    record.dataOffset,
    record.endOffset,
  );
  for (const child of children.records()) {
    if (child.type === PPT_RECORD.SLIDE_ATOM && child.length >= 16) {
      const view = new DataView(
        child.data.buffer,
        child.data.byteOffset,
        child.data.byteLength,
      );
      masterId = view.getUint32(12, true);
    }
    if (child.type === PPT_RECORD.PP_DRAWING) drawing = child.data;
  }

  const parsedDrawing = drawing
    ? await parsePptDrawing(drawing, theme, fonts, context)
    : undefined;
  const elements = parsedDrawing?.elements ?? [];
  const annotations = parsePptComments(
    record,
    `ppt-slide-${descriptor.persistId}`,
    descriptor.index,
    elements,
  );
  const transitionResult = parsePptTransition(record, descriptor.index);
  if (transitionResult.warning) {
    context.warnings.push(transitionResult.warning);
  }
  return {
    id: `ppt-slide-${descriptor.persistId}`,
    persistId: descriptor.persistId,
    slideId: descriptor.slideId,
    index: descriptor.index,
    width,
    height,
    masterId,
    hidden: descriptor.hidden,
    background: parsedDrawing?.background,
    annotations,
    transition: transitionResult.transition,
    warnings: transitionResult.warning ? [transitionResult.warning] : undefined,
    elements,
    sourceOffset: offset,
  };
}
