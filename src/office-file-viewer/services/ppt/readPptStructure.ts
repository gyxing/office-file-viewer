import {
  openCfbRandomAccess,
  type CfbDirectoryEntry,
  type CfbRandomAccessReader,
} from '../../shared/binary/cfb';
import { createBlobRandomAccessSource } from '../../shared/io/createBlobRandomAccessSource';
import { ResourceRegistry } from '../parsing/assembly/ResourceRegistry';
import {
  createPresentationPerformanceProfile,
  type PresentationPerformanceProfile,
} from '../presentation/presentationPerformance';
import type { PresentationWarning, ThemeModel } from '../presentation/types';
import { PPT_RECORD } from './binary/constants';
import { PptRecordReader } from './binary/PptRecordReader';
import { readPptEmbeddedChartsFromStream } from './chart/readEmbeddedCharts';
import {
  readPptDocumentBaseStructure,
  type PptDocumentReadStructure,
} from './document/readDocument';
import { readPptMaster } from './document/readMaster';
import type { PptNotesDescriptor } from './document/readSlideLists';
import { PptParseError } from './errors';
import { indexPptPictures } from './images';
import { buildPptEditChainFromStream } from './persistence';
import {
  createLocalPptEditChain,
  readPptPersistObject,
} from './readPptPersistObject';
import {
  createPptParseContext,
  type PptEditChain,
  type PptMasterModel,
  type PptParseContext,
  type PptSlideDescriptor,
} from './types';

export type PptArchiveProfile = {
  performance: PresentationPerformanceProfile;
  fileSize: number;
  mainStreamSize: number;
  resourceStreamSize: number;
};

export type ProfiledPptArchive = {
  reader: CfbRandomAccessReader;
  profile: PptArchiveProfile;
};

/** 大型 PPT Source 常驻的结构和共享解析上下文。 */
export type PptStructure = {
  reader: CfbRandomAccessReader;
  documentStream: NonNullable<ReturnType<CfbRandomAccessReader['openStream']>>;
  editChain: PptEditChain;
  parseContext: PptParseContext;
  resources: ResourceRegistry;
  width: number;
  height: number;
  theme: ThemeModel;
  masters: Map<number, PptMasterModel>;
  fonts: Map<number, string>;
  slideDescriptors: readonly PptSlideDescriptor[];
  notesBySlideId: ReadonlyMap<number, PptNotesDescriptor>;
  warnings: PresentationWarning[];
};

function findStream(entries: readonly CfbDirectoryEntry[], name: string) {
  return entries.find(
    (entry) =>
      entry.objectType === 'stream' &&
      entry.name.toLowerCase() === name.toLowerCase(),
  );
}

/** 根据 CFB 目录和共享阈值决定 PPT 是否启用按页数据源。 */
export function createPptArchiveProfile(
  fileSize: number,
  entries: readonly CfbDirectoryEntry[],
): PptArchiveProfile {
  const mainStreamSize =
    findStream(entries, 'PowerPoint Document')?.streamSize ?? 0;
  const resourceStreamSize = findStream(entries, 'Pictures')?.streamSize ?? 0;
  return {
    performance: createPresentationPerformanceProfile({
      slideCount: 0,
      cfbFileBytes: fileSize,
      cfbMainStreamBytes: mainStreamSize,
      cfbResourceStreamBytes: resourceStreamSize,
    }),
    fileSize,
    mainStreamSize,
    resourceStreamSize,
  };
}

/** 使用随机访问 CFB Reader 读取目录画像，不预先物化任何业务流。 */
export async function profilePptArchive(
  file: File,
  signal?: AbortSignal,
): Promise<ProfiledPptArchive> {
  const reader = await openCfbRandomAccess(
    createBlobRandomAccessSource(file),
    signal,
  );
  try {
    return {
      reader,
      profile: createPptArchiveProfile(file.size, reader.entries),
    };
  } catch (error) {
    await reader.close();
    throw error;
  }
}

function readNotesSlideId(
  documentStream: Uint8Array,
  editChain: PptEditChain,
  descriptor: PptNotesDescriptor,
) {
  const offset = editChain.persistOffsets.get(descriptor.persistId);
  if (offset === undefined) return undefined;
  const record = new PptRecordReader(
    documentStream,
    offset,
    documentStream.length,
  ).readRecord();
  if (!record || record.type !== PPT_RECORD.NOTES) return undefined;
  for (const child of new PptRecordReader(
    documentStream,
    record.dataOffset,
    record.endOffset,
  ).records()) {
    if (child.type !== PPT_RECORD.NOTES_ATOM || child.length < 4) continue;
    return new DataView(
      child.data.buffer,
      child.data.byteOffset,
      child.data.byteLength,
    ).getUint32(0, true);
  }
  return undefined;
}

/** 按 Persist Offset 读取 PPT 共享结构，不物化完整主流或正式 Slide。 */
export async function readPptStructure(
  archive: ProfiledPptArchive,
  signal?: AbortSignal,
): Promise<PptStructure> {
  const mainStream = archive.reader.openStream('PowerPoint Document');
  const currentUserStream = archive.reader.openStream('Current User');
  if (!mainStream || !currentUserStream) {
    throw new PptParseError(
      'PPT_REQUIRED_STREAM_MISSING',
      'PPT 文件缺少 PowerPoint Document 或 Current User 数据流',
    );
  }
  const context = createPptParseContext(async () => {
    if (signal?.aborted) {
      const error = new Error('PPT 解析已取消');
      error.name = 'AbortError';
      throw error;
    }
    await new Promise<void>((resolve) => {
      setTimeout(resolve, 0);
    });
  });
  const currentUser = await currentUserStream.materialize(signal);
  const editChain = await buildPptEditChainFromStream(
    mainStream,
    currentUser,
    context,
    signal,
  );
  const documentRecord = await readPptPersistObject(
    mainStream,
    editChain,
    editChain.documentPersistId,
    signal,
  );
  if (!documentRecord) {
    throw new PptParseError(
      'PPT_DOCUMENT_MISSING',
      '持久化目录中缺少根文档对象',
    );
  }
  const baseStructure = readPptDocumentBaseStructure(
    documentRecord.bytes,
    0,
    context,
  );
  const pictures = archive.reader.openStream('Pictures');
  if (pictures) {
    await indexPptPictures(pictures, context, signal);
  }
  await readPptEmbeddedChartsFromStream(
    documentRecord.bytes,
    mainStream,
    editChain,
    context,
    signal,
  );
  const masters = new Map<number, PptMasterModel>();
  for (const descriptor of baseStructure.descriptors.masters) {
    const record = await readPptPersistObject(
      mainStream,
      editChain,
      descriptor.persistId,
      signal,
    );
    if (!record) continue;
    const master = readPptMaster(
      record.bytes,
      createLocalPptEditChain(editChain, descriptor.persistId),
      descriptor,
      baseStructure.theme,
      baseStructure.fonts,
      context,
    );
    if (master) masters.set(master.id, master);
  }
  const documentStructure: PptDocumentReadStructure = {
    ...baseStructure,
    masters,
  };
  const notesBySlideId = new Map<number, PptNotesDescriptor>();
  for (const descriptor of documentStructure.descriptors.notes) {
    const record = await readPptPersistObject(
      mainStream,
      editChain,
      descriptor.persistId,
      signal,
    );
    if (!record) continue;
    const slideId = readNotesSlideId(
      record.bytes,
      createLocalPptEditChain(editChain, descriptor.persistId),
      descriptor,
    );
    if (slideId) notesBySlideId.set(slideId, descriptor);
  }
  const resources = new ResourceRegistry();
  for (const resource of context.resources.splice(0)) {
    await resources.register(resource);
  }

  return {
    reader: archive.reader,
    documentStream: mainStream,
    editChain,
    parseContext: context,
    resources,
    width: documentStructure.width,
    height: documentStructure.height,
    theme: documentStructure.theme,
    masters: documentStructure.masters,
    fonts: documentStructure.fonts,
    slideDescriptors: documentStructure.descriptors.slides,
    notesBySlideId,
    warnings: context.warnings,
  };
}
