import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import { imageMimeType } from '../../shared/ooxml/media';
import type { OfficeArchiveReader } from '../../shared/ooxml/OfficeArchiveReader';
import type { OfficeResourceSource } from '../resource-store/types';
import {
  createDocxParseContext,
  type DocxParseContext,
} from './docxParsingContext';

/** DOCX 压缩包内共享关系、主题和资源的解析上下文。 */
export type DocxPackageContext = {
  /** 压缩包或复合文档包含的条目。 */
  entries: OfficeEntryMap;
  /** 当前解析任务共享的上下文。 */
  parseContext: DocxParseContext;
};

/** 读取主文档以外的小型 XML，并为媒体建立不解压的懒资源引用。 */
export async function loadDocxPackageContext(
  reader: OfficeArchiveReader,
  bodyNode: Element,
  resourceNamespace: string,
  signal?: AbortSignal,
): Promise<DocxPackageContext> {
  const xmlEntries = reader
    .list()
    .filter(
      (entry) =>
        entry.path !== 'word/document.xml' &&
        (entry.path.endsWith('.xml') || entry.path.endsWith('.rels')),
    );
  const entries: OfficeEntryMap = new Map();
  let nextIndex = 0;
  const readNext = async () => {
    while (nextIndex < xmlEntries.length) {
      if (signal?.aborted) throw signal.reason;
      const entry = xmlEntries[nextIndex];
      nextIndex += 1;
      entries.set(entry.path, await reader.readText(entry.path, signal));
    }
  };
  await Promise.all(
    Array.from({ length: Math.min(4, Math.max(1, xmlEntries.length)) }, () =>
      readNext(),
    ),
  );

  const byPath: Record<string, OfficeResourceSource> = {};
  const byName: Record<string, OfficeResourceSource> = {};
  reader.list('word/media/').forEach((entry) => {
    const mimeType = imageMimeType(entry.path);
    const source: OfficeResourceSource = {
      kind: 'lazy',
      id: `${resourceNamespace}:${entry.path}`,
      mimeType,
      size: entry.uncompressedSize,
      load: (resourceSignal) =>
        reader.readBlob(entry.path, mimeType, resourceSignal),
    };
    byPath[entry.path] = source;
    byName[entry.path.split('/').pop() ?? entry.path] = source;
  });

  return {
    entries,
    parseContext: createDocxParseContext(entries, {
      bodyNode,
      media: { byPath, byName },
    }),
  };
}
