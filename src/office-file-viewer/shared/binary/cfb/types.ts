/** CFB 目录项支持的对象种类。 */
export type CfbObjectType = 'storage' | 'stream' | 'root';

/** CFB 复合文档目录中的单个存储或数据流。 */
export type CfbDirectoryEntry = {
  /** 在所属集合中的唯一标识。 */
  id: number;
  /** 面向用户展示的名称。 */
  name: string;
  /** 在压缩包、复合文档或资源表中的路径。 */
  path: string;
  /** 目录项是存储、数据流还是根存储。 */
  objectType: CfbObjectType;
  /** 在 CFB 扇区链中的扇区索引。 */
  startSector: number;
  /** 占用或消费的字节数。 */
  streamSize: number;
  /** 目录红黑树中的左兄弟项标识。 */
  leftSiblingId: number;
  /** 目录红黑树中的右兄弟项标识。 */
  rightSiblingId: number;
  /** 存储目录树中的首个子项标识。 */
  childId: number;
};

/** 完成物化的 CFB 目录和命名数据流。 */
export type CfbFile = {
  /** 按目录编号排列的 CFB 目录项。 */
  entries: CfbDirectoryEntry[];
  /** 按流名称索引的 CFB 数据流映射。 */
  streams: Map<string, Uint8Array>;
  /** 按候选名称读取首个匹配的数据流。 */
  getStream: (...names: string[]) => Uint8Array | undefined;
  /** 判断复合文档是否包含指定目录项。 */
  hasEntry: (name: string) => boolean;
};

/** 读取 CFB 结构和数据流时使用的兼容与取消选项。 */
export type CfbReadOptions = {
  /** 长任务需要让出主线程时执行的回调。 */
  yieldIfNeeded?: () => Promise<void>;
  /** 兼容省略最后扇区零填充的 CFB 生成器；默认保持严格校验。 */
  allowPartialFinalSector?: boolean;
  /** 取消 CFB 结构或流读取任务的统一信号。 */
  signal?: AbortSignal;
};

/** 提供单个 CFB 流的有界随机读取和兼容物化能力。 */
export interface CfbStreamReader {
  /** 当前处理的压缩包或复合文档条目。 */
  readonly entry: CfbDirectoryEntry;
  /** 读取指定偏移与长度的连续字节。 */
  read(
    offset: number,
    length: number,
    signal?: AbortSignal,
  ): Promise<Uint8Array>;
  /** 将当前数据流完整读取为内存字节。 */
  materialize(signal?: AbortSignal): Promise<Uint8Array>;
}

/** 提供 CFB 目录索引和命名流随机读取能力。 */
export interface CfbRandomAccessReader {
  /** 压缩包或复合文档包含的条目。 */
  readonly entries: readonly CfbDirectoryEntry[];
  /** 判断 CFB 复合文档内是否存在指定条目。 */
  hasEntry(name: string): boolean;
  /** 打开指定条目的随机读取数据流。 */
  openStream(...names: string[]): CfbStreamReader | undefined;
  /** 幂等关闭读取器并释放底层数据源。 */
  close(): Promise<void>;
}
