import type { OfficeArchiveResourcePolicy } from '../resource/OfficeResourcePolicy';
import type { OfficeEntryMap } from './archive';

/** OOXML 归档条目的路径和压缩体积。 */
export type OfficeArchiveEntry = {
  /** 在压缩包、复合文档或资源表中的路径。 */
  path: string;
  /** 压缩包中相关内容的压缩大小，单位为字节。 */
  compressedSize: number;
  /** 相关内容解压后的大小，单位为字节。 */
  uncompressedSize: number;
};

/** 描述打开 OOXML 归档时可用的取消选项。 */
export type OfficeArchiveOpenOptions = {
  /** 用于取消当前异步操作的信号。 */
  signal?: AbortSignal;
  /** 宿主可选配置的归档读取上限；未提供时不限制。 */
  resourcePolicy?: OfficeArchiveResourcePolicy;
};

/**
 * 为普通物化解析和后续大文件懒读取提供统一的 OOXML 归档访问协议。
 */
export interface OfficeArchiveReader {
  /** 判断压缩包内是否存在指定条目。 */
  has(path: string): boolean;
  /** 列出压缩包内全部条目路径。 */
  list(prefix?: string): OfficeArchiveEntry[];
  /** 以 UTF-8 文本读取指定压缩包条目。 */
  readText(path: string, signal?: AbortSignal): Promise<string>;
  /** 以字节数组读取指定压缩包条目。 */
  readBinary(path: string, signal?: AbortSignal): Promise<Uint8Array>;
  /** 以 Blob 读取指定压缩包条目。 */
  readBlob(
    path: string,
    mimeType?: string,
    signal?: AbortSignal,
  ): Promise<Blob>;
  /** 打开指定条目的随机读取数据流。 */
  openStream(
    path: string,
    signal?: AbortSignal,
  ): Promise<ReadableStream<Uint8Array>>;
  /** 将当前数据流完整读取为内存字节。 */
  materialize(signal?: AbortSignal): Promise<OfficeEntryMap>;
  /** 幂等关闭读取器并释放底层数据源。 */
  close(): Promise<void>;
}
