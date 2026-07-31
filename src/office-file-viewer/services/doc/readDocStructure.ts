/** 读取 DOC FIB 基础结构所需的最少字节数。 */
const DOC_FIB_BASE_MIN_BYTES = 12;
/** DOC FIB 标志字段相对流起点的字节偏移。 */
const DOC_FIB_FLAGS_OFFSET = 10;
/** 用于选择 DOC 0Table 或 1Table 流的 FIB 位掩码。 */
const DOC_FIB_WHICH_TABLE_MASK = 0x0200;

/** 从 FIB Base 读取当前文档使用的 Table 流名称。 */
export function readDocTableStreamName(
  fibBase: Uint8Array,
): '0Table' | '1Table' {
  if (fibBase.length < DOC_FIB_BASE_MIN_BYTES) {
    throw new Error('DOC WordDocument 流缺少完整 FIB Base');
  }
  const view = new DataView(
    fibBase.buffer,
    fibBase.byteOffset,
    fibBase.byteLength,
  );
  return view.getUint16(DOC_FIB_FLAGS_OFFSET, true) & DOC_FIB_WHICH_TABLE_MASK
    ? '1Table'
    : '0Table';
}
