import type { Entry, FileEntry, ZipReader } from '@zip.js/zip.js';

import { OfficeResourceLimitError } from '../resource/OfficeResourceLimitError';
import {
  validateOfficeResourcePolicy,
  type OfficeArchiveResourcePolicy,
} from '../resource/OfficeResourcePolicy';
import type {
  OfficeArchiveEntry,
  OfficeArchiveReader,
} from './OfficeArchiveReader';
import type { OfficeEntryMap, OfficeZipInput } from './archive';

/** 运行时加载的 zip.js 核心模块接口。 */
type ZipJsCoreModule = typeof import('@zip.js/zip.js/lib/zip-core-native.js');
/** 支持批量枚举条目的 zip.js 读取器。 */
type ZipReaderWithEntries = Pick<ZipReader<unknown>, 'close' | 'getEntries'>;

/** zip.js 条目名称与读取句柄的索引记录。 */
type ArchiveRecord = {
  /** 归档条目的规范化路径和压缩体积。 */
  metadata: OfficeArchiveEntry;
  /** 当前处理的压缩包或复合文档条目。 */
  entry: FileEntry;
};

/** Office 压缩包条目并发读取的最大任务数。 */
const OFFICE_ENTRY_READ_CONCURRENCY = 4;
// 组件库无法预知消费者的 Worker 资源发布路径，条目解压统一在当前执行上下文运行。
/** zip.js 读取 Office 压缩包条目时共享的安全选项。 */
const OFFICE_ENTRY_READ_OPTIONS = {
  checkOverlappingEntry: true,
  useWebWorkers: false,
} as const;

/** 在浏览器环境识别 File/Blob，避免服务端模块求值访问 Blob 构造器。 */
function isBlobInput(input: OfficeZipInput): input is Blob {
  return typeof Blob !== 'undefined' && input instanceof Blob;
}

/** 统一 ZIP 内部路径，避免不同分隔符绕过重复条目检查。 */
function normalizeArchivePath(path: string) {
  return path.replace(/\\/g, '/').replace(/^\/+/, '');
}

/** 在进入耗时解压前保留 AbortError 或调用方自定义的取消原因。 */
function throwIfAborted(signal?: AbortSignal) {
  if (!signal?.aborted) {
    return;
  }
  if (signal.reason !== undefined) {
    throw signal.reason;
  }
  const error = new Error('Office 包读取已取消');
  error.name = 'AbortError';
  throw error;
}

/** 将外部取消信号转发给当前归档读取任务，并返回事件清理函数。 */
function linkAbortSignal(
  signal: AbortSignal | undefined,
  controller: AbortController,
) {
  if (!signal) {
    return () => undefined;
  }
  const abort = () => controller.abort(signal.reason);
  if (signal.aborted) {
    abort();
    return () => undefined;
  }
  signal.addEventListener('abort', abort, { once: true });
  return () => signal.removeEventListener('abort', abort);
}

/** 判断条目是否沿用旧归档层的 UTF-8 字符串读取语义。 */
function isTextEntry(path: string) {
  return /\.xml$/i.test(path) || /\.rels$/i.test(path);
}

/** 把底层异常转换为包含条目路径的中文错误。 */
function createEntryError(path: string, error: unknown) {
  const message = error instanceof Error ? error.message : '未知错误';
  return new Error(`Office 包条目解压失败（${path}）：${message}`);
}

/** 计算压缩比；空条目不会被视为高压缩比内容。 */
function getCompressionRatio(compressedSize: number, uncompressedSize: number) {
  if (uncompressedSize <= 0) return 0;
  if (compressedSize <= 0) return Number.POSITIVE_INFINITY;
  return uncompressedSize / compressedSize;
}

/** 在解压前使用中央目录元数据执行宿主配置的归档资源检查。 */
function assertArchiveWithinPolicy(
  entries: readonly Entry[],
  policy?: OfficeArchiveResourcePolicy,
) {
  validateOfficeResourcePolicy(policy);
  if (!policy) return;
  const fileEntries = entries.filter((entry) => !entry.directory);
  if (
    policy.maxArchiveEntries !== undefined &&
    fileEntries.length > policy.maxArchiveEntries
  ) {
    throw new OfficeResourceLimitError(
      'ARCHIVE_ENTRY_COUNT_LIMIT_EXCEEDED',
      'Office 包条目数量超过宿主配置的上限',
      { limit: policy.maxArchiveEntries, actual: fileEntries.length },
    );
  }

  let compressedTotal = 0;
  let uncompressedTotal = 0;
  for (const entry of fileEntries) {
    compressedTotal += entry.compressedSize;
    uncompressedTotal += entry.uncompressedSize;
    const path = normalizeArchivePath(entry.filename);
    if (
      policy.maxArchiveEntryBytes !== undefined &&
      entry.uncompressedSize > policy.maxArchiveEntryBytes
    ) {
      throw new OfficeResourceLimitError(
        'ARCHIVE_ENTRY_SIZE_LIMIT_EXCEEDED',
        `Office 包条目超过宿主配置的解压上限（${path}）`,
        {
          limit: policy.maxArchiveEntryBytes,
          actual: entry.uncompressedSize,
          path,
        },
      );
    }
    const ratio = getCompressionRatio(
      entry.compressedSize,
      entry.uncompressedSize,
    );
    if (
      policy.maxCompressionRatio !== undefined &&
      ratio > policy.maxCompressionRatio
    ) {
      throw new OfficeResourceLimitError(
        'ARCHIVE_COMPRESSION_RATIO_LIMIT_EXCEEDED',
        `Office 包条目压缩比超过宿主配置的上限（${path}）`,
        { limit: policy.maxCompressionRatio, actual: ratio, path },
      );
    }
  }

  if (
    policy.maxArchiveInflatedBytes !== undefined &&
    uncompressedTotal > policy.maxArchiveInflatedBytes
  ) {
    throw new OfficeResourceLimitError(
      'ARCHIVE_TOTAL_SIZE_LIMIT_EXCEEDED',
      'Office 包累计解压大小超过宿主配置的上限',
      { limit: policy.maxArchiveInflatedBytes, actual: uncompressedTotal },
    );
  }
  const totalRatio = getCompressionRatio(compressedTotal, uncompressedTotal);
  if (
    policy.maxCompressionRatio !== undefined &&
    totalRatio > policy.maxCompressionRatio
  ) {
    throw new OfficeResourceLimitError(
      'ARCHIVE_COMPRESSION_RATIO_LIMIT_EXCEEDED',
      'Office 包整体压缩比超过宿主配置的上限',
      { limit: policy.maxCompressionRatio, actual: totalRatio },
    );
  }
}
/**
 * 使用 zip.js 读取 OOXML 归档，并保持现有 OfficeEntryMap 的物化结果。
 */
export class ZipJsOfficeArchiveReader implements OfficeArchiveReader {
  private readonly entriesByPath = new Map<string, ArchiveRecord>();

  private readonly entriesInOrder: ArchiveRecord[] = [];

  private closePromise?: Promise<void>;

  constructor(
    private readonly zipModule: ZipJsCoreModule,
    private readonly zipReader: ZipReaderWithEntries,
    private readonly useCompressionStream: boolean,
    entries: Entry[],
    resourcePolicy?: OfficeArchiveResourcePolicy,
  ) {
    assertArchiveWithinPolicy(entries, resourcePolicy);
    for (const entry of entries) {
      if (entry.directory) {
        continue;
      }
      const path = normalizeArchivePath(entry.filename);
      if (this.entriesByPath.has(path)) {
        throw new Error(`Office 包包含重复条目路径（${path}）`);
      }
      const record: ArchiveRecord = {
        entry,
        metadata: {
          path,
          compressedSize: entry.compressedSize,
          uncompressedSize: entry.uncompressedSize,
        },
      };
      this.entriesByPath.set(path, record);
      this.entriesInOrder.push(record);
    }
  }

  has(path: string) {
    return this.entriesByPath.has(normalizeArchivePath(path));
  }

  list(prefix = '') {
    const normalizedPrefix = normalizeArchivePath(prefix);
    return this.entriesInOrder
      .filter(({ metadata }) => metadata.path.startsWith(normalizedPrefix))
      .map(({ metadata }) => ({ ...metadata }));
  }

  async readText(path: string, signal?: AbortSignal) {
    const record = this.getRecord(path);
    throwIfAborted(signal);
    return record.entry.getData(new this.zipModule.TextWriter(), {
      ...OFFICE_ENTRY_READ_OPTIONS,
      signal,
      useCompressionStream: this.useCompressionStream,
    });
  }

  async readBinary(path: string, signal?: AbortSignal) {
    const record = this.getRecord(path);
    throwIfAborted(signal);
    return record.entry.getData(new this.zipModule.Uint8ArrayWriter(), {
      ...OFFICE_ENTRY_READ_OPTIONS,
      signal,
      useCompressionStream: this.useCompressionStream,
    });
  }

  async readBlob(path: string, mimeType = '', signal?: AbortSignal) {
    const record = this.getRecord(path);
    throwIfAborted(signal);
    return record.entry.getData(new this.zipModule.BlobWriter(mimeType), {
      ...OFFICE_ENTRY_READ_OPTIONS,
      signal,
      useCompressionStream: this.useCompressionStream,
    });
  }

  async openStream(path: string, signal?: AbortSignal) {
    const record = this.getRecord(path);
    throwIfAborted(signal);
    const taskController = new AbortController();
    const unlinkSignal = linkAbortSignal(signal, taskController);
    const transport = new TransformStream<Uint8Array>();
    const transportReader = transport.readable.getReader();
    let readerReleased = false;

    const releaseReader = () => {
      if (!readerReleased) {
        transportReader.releaseLock();
        readerReleased = true;
      }
    };

    const completion = record.entry.getData(transport.writable, {
      ...OFFICE_ENTRY_READ_OPTIONS,
      signal: taskController.signal,
      useCompressionStream: this.useCompressionStream,
    });
    void completion
      .catch((error) => transport.writable.abort(error))
      .catch(() => undefined)
      .finally(unlinkSignal);

    return new ReadableStream<Uint8Array>({
      async pull(controller) {
        try {
          const result = await transportReader.read();
          if (result.done) {
            releaseReader();
            controller.close();
            return;
          }
          controller.enqueue(result.value);
        } catch (error) {
          releaseReader();
          controller.error(error);
        }
      },
      async cancel(reason) {
        taskController.abort(reason);
        try {
          await transportReader.cancel(reason);
        } finally {
          releaseReader();
          unlinkSignal();
        }
      },
    });
  }

  async materialize(signal?: AbortSignal): Promise<OfficeEntryMap> {
    throwIfAborted(signal);
    const taskController = new AbortController();
    const unlinkSignal = linkAbortSignal(signal, taskController);
    const results = new Array<
      readonly [path: string, data: string | Uint8Array]
    >(this.entriesInOrder.length);
    let nextIndex = 0;
    let primaryError: Error | undefined;

    const readNextEntry = async () => {
      while (
        !taskController.signal.aborted &&
        nextIndex < this.entriesInOrder.length
      ) {
        const entryIndex = nextIndex;
        nextIndex += 1;
        const record = this.entriesInOrder[entryIndex];
        const path = record.metadata.path;
        try {
          const data = isTextEntry(path)
            ? await this.readText(path, taskController.signal)
            : await this.readBinary(path, taskController.signal);
          results[entryIndex] = [path, data];
        } catch (error) {
          if (signal?.aborted) {
            throwIfAborted(signal);
          }
          if (!primaryError) {
            primaryError =
              error instanceof Error && error.name === 'AbortError'
                ? error
                : createEntryError(path, error);
            taskController.abort(primaryError);
          }
          throw primaryError;
        }
      }
      if (primaryError) {
        throw primaryError;
      }
      throwIfAborted(signal);
    };

    try {
      const workerCount = Math.min(
        OFFICE_ENTRY_READ_CONCURRENCY,
        this.entriesInOrder.length,
      );
      await Promise.all(
        Array.from({ length: workerCount }, () => readNextEntry()),
      );
      return new Map(results);
    } finally {
      unlinkSignal();
    }
  }

  close() {
    this.closePromise ??= this.zipReader.close();
    return this.closePromise;
  }

  private getRecord(path: string) {
    if (this.closePromise) {
      throw new Error('Office 包读取器已关闭');
    }
    const normalizedPath = normalizeArchivePath(path);
    const record = this.entriesByPath.get(normalizedPath);
    if (!record) {
      throw new Error(`Office 包中不存在条目（${normalizedPath}）`);
    }
    return record;
  }
}

/**
 * 用已动态加载的 zip.js 核心打开归档，并在构造失败时立即释放 Reader。
 */
export async function createZipJsOfficeArchiveReader(
  input: OfficeZipInput,
  zipModule: ZipJsCoreModule,
  useCompressionStream: boolean,
  signal?: AbortSignal,
  resourcePolicy?: OfficeArchiveResourcePolicy,
) {
  throwIfAborted(signal);
  const zipReader = isBlobInput(input)
    ? new zipModule.ZipReader(new zipModule.BlobReader(input))
    : new zipModule.ZipReader(
        new zipModule.Uint8ArrayReader(
          input instanceof Uint8Array ? input : new Uint8Array(input),
        ),
      );

  try {
    const entries = await zipReader.getEntries();
    throwIfAborted(signal);
    return new ZipJsOfficeArchiveReader(
      zipModule,
      zipReader,
      useCompressionStream,
      entries,
      resourcePolicy,
    );
  } catch (error) {
    try {
      await zipReader.close();
    } catch {
      // 打开阶段优先保留中央目录或路径校验产生的原始错误。
    }
    throw error;
  }
}
