/** DOC 二进制按需数据源。 */
export interface DocBinarySource {
  /** 原始 CFB 文件字节数。 */
  readonly fileSize: number;
  /** WordDocument 流字节数。 */
  readonly wordDocumentSize: number;
  /** FIB 指定的 0Table 或 1Table 流名称。 */
  readonly tableStreamName: '0Table' | '1Table';
  /** FIB 指定的 Table 流字节数。 */
  readonly tableSize: number;
  /** 可选 Data 流字节数。 */
  readonly dataSize: number;
  /** 按范围读取 WordDocument，避免先物化整个 CFB 文件。 */
  readWordDocument(
    offset: number,
    length: number,
    signal?: AbortSignal,
  ): Promise<Uint8Array>;
  /** 按范围读取 FIB 指定的 Table 流。 */
  readTable(
    offset: number,
    length: number,
    signal?: AbortSignal,
  ): Promise<Uint8Array>;
  /** 按范围读取图片和 OfficeArt 常用的 Data 流。 */
  readData(
    offset: number,
    length: number,
    signal?: AbortSignal,
  ): Promise<Uint8Array>;
  /** 读取 ObjectPool 中嵌入对象的静态预览流，并保持 CFB 目录顺序。 */
  readObjectPreviewStreams(
    signal?: AbortSignal,
  ): Promise<Array<readonly [string, Uint8Array]>>;
  /** 幂等释放 CFB Reader 和底层 Blob 数据源。 */
  dispose(): Promise<void>;
}
