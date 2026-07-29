import type { OfficeEntryMap } from './archive';

/** 描述 OOXML 归档条目的规范化路径和压缩体积。 */
export type OfficeArchiveEntry = {
  path: string;
  compressedSize: number;
  uncompressedSize: number;
};

/** 描述打开 OOXML 归档时可用的取消选项。 */
export type OfficeArchiveOpenOptions = {
  signal?: AbortSignal;
};

/**
 * 为普通物化解析和后续大文件懒读取提供统一的 OOXML 归档访问协议。
 */
export interface OfficeArchiveReader {
  has(path: string): boolean;
  list(prefix?: string): OfficeArchiveEntry[];
  readText(path: string, signal?: AbortSignal): Promise<string>;
  readBinary(path: string, signal?: AbortSignal): Promise<Uint8Array>;
  readBlob(
    path: string,
    mimeType?: string,
    signal?: AbortSignal,
  ): Promise<Blob>;
  openStream(
    path: string,
    signal?: AbortSignal,
  ): Promise<ReadableStream<Uint8Array>>;
  materialize(signal?: AbortSignal): Promise<OfficeEntryMap>;
  close(): Promise<void>;
}
