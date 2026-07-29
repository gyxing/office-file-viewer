/** 描述 CfbObjectType 在 CFB 复合文档中的数据结构。 */
export type CfbObjectType = 'storage' | 'stream' | 'root';

/** 表示CFB 复合文档集合中的一个条目。 */
export type CfbDirectoryEntry = {
  /** CfbDirectoryEntry 在所属文档或任务中的唯一标识。 */
  id: number;
  /** CfbDirectoryEntry 的可读名称。 */
  name: string;
  /** CfbDirectoryEntry 在压缩包、复合文档或图形数据中的路径。 */
  path: string;
  /** CfbDirectoryEntry 关联的 objectType 结构；字段形状由 CfbObjectType 定义。 */
  objectType: CfbObjectType;
  /** CfbDirectoryEntry 在 CFB 扇区链中的扇区索引。 */
  startSector: number;
  /** CfbDirectoryEntry 占用或消费的字节数。 */
  streamSize: number;
  /** CfbDirectoryEntry 在源文件记录中的数字标识。 */
  leftSiblingId: number;
  /** CfbDirectoryEntry 在源文件记录中的数字标识。 */
  rightSiblingId: number;
  /** CfbDirectoryEntry 在源文件记录中的数字标识。 */
  childId: number;
};

/** 描述 CfbFile 在 CFB 复合文档中的数据结构。 */
export type CfbFile = {
  /** CfbFile 包含的 entries 有序集合。 */
  entries: CfbDirectoryEntry[];
  /** CfbFile 按流名称索引的 CFB 数据流映射。 */
  streams: Map<string, Uint8Array>;
  /** CfbFile 执行 getStream 操作时调用的函数。 */
  getStream: (...names: string[]) => Uint8Array | undefined;
  /** CfbFile 执行 hasEntry 操作时调用的函数。 */
  hasEntry: (name: string) => boolean;
};

/** 定义CFB 复合文档的可选配置。 */
export type CfbReadOptions = {
  /** CfbReadOptions 执行 yieldIfNeeded 操作时调用的函数。 */
  yieldIfNeeded?: () => Promise<void>;
  /** 兼容省略最后扇区零填充的 CFB 生成器；默认保持严格校验。 */
  allowPartialFinalSector?: boolean;
  /** 取消 CFB 结构或流读取任务的统一信号。 */
  signal?: AbortSignal;
};

/** 提供单个 CFB 流的有界随机读取和兼容物化能力。 */
export interface CfbStreamReader {
  readonly entry: CfbDirectoryEntry;
  read(
    offset: number,
    length: number,
    signal?: AbortSignal,
  ): Promise<Uint8Array>;
  materialize(signal?: AbortSignal): Promise<Uint8Array>;
}

/** 提供 CFB 目录索引和命名流随机读取能力。 */
export interface CfbRandomAccessReader {
  readonly entries: readonly CfbDirectoryEntry[];
  hasEntry(name: string): boolean;
  openStream(...names: string[]): CfbStreamReader | undefined;
  close(): Promise<void>;
}
