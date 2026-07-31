/** 提供对 Blob、文件或其他容器的有界随机读取能力。 */
export interface RandomAccessSource {
  /** 数据源的完整字节长度。 */
  readonly size: number;
  /** 读取指定偏移和长度的连续字节。 */
  read(
    offset: number,
    length: number,
    signal?: AbortSignal,
  ): Promise<Uint8Array>;
  /** 幂等关闭数据源，关闭后不再接受读取。 */
  close(): Promise<void>;
}

/** 随机读取失败时使用的稳定错误结构。 */
export type RandomAccessSourceError = Error & {
  /** 供程序识别当前情况的稳定代码。 */
  code: 'INVALID_RANDOM_ACCESS_RANGE' | 'RANDOM_ACCESS_SOURCE_CLOSED';
};
