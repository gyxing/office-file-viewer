import { createMaterializedPresentationSource } from './createMaterializedPresentationSource';
import type { PresentationSource } from './PresentationSource';
import type { PresentationDocument } from './types';

const sourceByDocument = new WeakMap<
  PresentationDocument,
  PresentationSource
>();

/** 为物化文稿注册其实际数据源，避免渲染层依赖解析器实现。 */
export function registerPresentationSource(
  document: PresentationDocument,
  source: PresentationSource,
) {
  sourceByDocument.set(document, source);
}

/** 获取文稿对应的数据源；普通模型按引用只创建一次轻量适配器。 */
export function getPresentationSource(document: PresentationDocument) {
  const existing = sourceByDocument.get(document);
  if (existing) return existing;
  const source = createMaterializedPresentationSource(document);
  sourceByDocument.set(document, source);
  return source;
}
