import { resolveSlideResources } from '../parsing/assembly/DocumentAssembler';
import type { SlideModel } from '../presentation/types';
import { adaptPptSlide } from './adapter';
import { readPptSlide } from './document/readSlide';
import {
  createLocalPptEditChain,
  readPptPersistObject,
} from './readPptPersistObject';
import type { PptStructure } from './readPptStructure';

/** 通过 Persist Directory 的目标偏移解析一页 PPT，并合并既有母版规则。 */
export async function loadPptSlide(
  structure: PptStructure,
  index: number,
  signal?: AbortSignal,
): Promise<SlideModel> {
  if (signal?.aborted) {
    const error = new Error('PPT 幻灯片读取已取消');
    error.name = 'AbortError';
    throw error;
  }
  const descriptor = structure.slideDescriptors[index];
  if (!descriptor) throw new RangeError(`幻灯片索引超出范围：${index}`);
  const record = await readPptPersistObject(
    structure.documentStream,
    structure.editChain,
    descriptor.persistId,
    signal,
  );
  if (!record) {
    throw new Error(`第 ${descriptor.index} 页缺少 PPT 持久化对象`);
  }
  const slide = readPptSlide(
    record.bytes,
    createLocalPptEditChain(structure.editChain, descriptor.persistId),
    descriptor,
    structure.width,
    structure.height,
    structure.theme,
    structure.fonts,
    structure.parseContext,
  );
  if (!slide) {
    throw new Error(`第 ${descriptor.index} 页不是有效的 PPT Slide`);
  }
  slide.sourceOffset = record.sourceOffset;
  slide.background ??= structure.masters.get(slide.masterId ?? Number.NaN)
    ?.background ?? {
    fill: structure.theme.colorScheme.lt1 ?? '#ffffff',
  };
  const result = adaptPptSlide(slide, structure.masters);
  resolveSlideResources(result, structure.resources);
  return result;
}
