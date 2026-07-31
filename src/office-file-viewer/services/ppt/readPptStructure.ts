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

/** PPT 复合文档的大文件判定指标。 */
export type PptArchiveProfile = {
  /** 当前文档的性能统计信息。 */
  performance: PresentationPerformanceProfile;
  /** 源文件总大小，单位为字节。 */
  fileSize: number;
  /** 主文档数据流大小，单位为字节。 */
  mainStreamSize: number;
  /** 图片等资源数据流大小，单位为字节。 */
  resourceStreamSize: number;
};

/** 附带性能档案的 PPT 复合文档读取器。 */
export type ProfiledPptArchive = {
  /** 用于按需读取源数据的读取器。 */
  reader: CfbRandomAccessReader;
  /** 控制解析或渲染策略的性能档案。 */
  profile: PptArchiveProfile;
};

/** 大型 PPT Source 常驻的结构和共享解析上下文。 */
export type PptStructure = {
  /** 用于按需读取源数据的读取器。 */
  reader: CfbRandomAccessReader;
  /** 按需读取 PowerPoint Document 主流的接口。 */
  documentStream: NonNullable<ReturnType<CfbRandomAccessReader['openStream']>>;
  /** 当前文档最近一次有效保存对应的编辑链。 */
  editChain: PptEditChain;
  /** 当前解析任务共享的上下文。 */
  parseContext: PptParseContext;
  /** 集中管理图片、字体和对象地址等可释放资源。 */
  resources: ResourceRegistry;
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 当前文档使用的主题颜色和字体配置。 */
  theme: ThemeModel;
  /** 按母版标识索引的 PPT 母版模型。 */
  masters: Map<number, PptMasterModel>;
  /** 按字体编号索引的字体族名称。 */
  fonts: Map<number, string>;
  /** 按演示顺序排列的幻灯片描述信息。 */
  slideDescriptors: readonly PptSlideDescriptor[];
  /** 按幻灯片标识索引的备注页描述信息。 */
  notesBySlideId: ReadonlyMap<number, PptNotesDescriptor>;
  /** 解析时产生但不阻止继续预览的警告。 */
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
