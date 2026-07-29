/** 共享容器读取模式的切换阈值；这些数值不是业务数量限制。 */
export const OFFICE_LARGE_FILE_THRESHOLDS = {
  ooxmlCompressedBytes: 12 * 1024 * 1024,
  ooxmlUncompressedBytes: 64 * 1024 * 1024,
  ooxmlMainXmlBytes: 8 * 1024 * 1024,
  ooxmlSingleMediaBytes: 32 * 1024 * 1024,
  cfbFileBytes: 32 * 1024 * 1024,
  cfbMainStreamBytes: 24 * 1024 * 1024,
  cfbResourceStreamBytes: 32 * 1024 * 1024,
  slowTaskMilliseconds: 50,
} as const;
