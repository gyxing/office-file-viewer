/** 浏览器回退字体仅保留字形、字符映射和基础度量所需的 SFNT 表。 */
const BROWSER_SAFE_SFNT_TABLES = new Set([
  'CFF ',
  'CFF2',
  'OS/2',
  'cmap',
  'cvt ',
  'fpgm',
  'gasp',
  'glyf',
  'head',
  'hhea',
  'hmtx',
  'loca',
  'maxp',
  'name',
  'post',
  'prep',
]);

/** OpenType 整体校验和要求达到的固定魔数。 */
const SFNT_CHECKSUM_MAGIC = 0xb1b0afba;

/** SFNT 目录中的单个字体表。 */
type SfntTable = {
  /** 四字节表标签。 */
  tag: string;
  /** 原字体中的字节偏移。 */
  offset: number;
  /** 表的实际字节长度。 */
  length: number;
};

function alignFour(value: number) {
  return (value + 3) & ~3;
}

function readTag(bytes: Uint8Array, offset: number) {
  return String.fromCharCode(
    bytes[offset],
    bytes[offset + 1],
    bytes[offset + 2],
    bytes[offset + 3],
  );
}

function writeTag(bytes: Uint8Array, offset: number, tag: string) {
  for (let index = 0; index < 4; index += 1) {
    bytes[offset + index] = tag.charCodeAt(index);
  }
}

/** 按 4 字节大端整数累加 OpenType 表或完整字体校验和。 */
function calculateChecksum(bytes: Uint8Array) {
  const padded = new Uint8Array(alignFour(bytes.length));
  padded.set(bytes);
  const view = new DataView(padded.buffer);
  let checksum = 0;
  for (let offset = 0; offset < padded.length; offset += 4) {
    checksum = (checksum + view.getUint32(offset)) >>> 0;
  }
  return checksum;
}

/** 丢弃旧字体中浏览器会严格拒绝的非 BMP cmap，只保留有效格式 4。 */
function createBmpCmap(bytes: Uint8Array) {
  if (bytes.length < 4) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const recordCount = view.getUint16(2);
  if (4 + recordCount * 8 > bytes.length) return undefined;
  let candidateOffset: number | undefined;
  for (let index = 0; index < recordCount; index += 1) {
    const recordOffset = 4 + index * 8;
    const platform = view.getUint16(recordOffset);
    const encoding = view.getUint16(recordOffset + 2);
    const subtableOffset = view.getUint32(recordOffset + 4);
    if (
      subtableOffset + 4 > bytes.length ||
      view.getUint16(subtableOffset) !== 4
    ) {
      continue;
    }
    candidateOffset ??= subtableOffset;
    if (platform === 3 && encoding === 1) {
      candidateOffset = subtableOffset;
      break;
    }
  }
  if (candidateOffset === undefined) return undefined;
  const subtableLength = view.getUint16(candidateOffset + 2);
  if (candidateOffset + subtableLength > bytes.length) return undefined;
  const output = new Uint8Array(12 + subtableLength);
  const outputView = new DataView(output.buffer);
  outputView.setUint16(0, 0);
  outputView.setUint16(2, 1);
  outputView.setUint16(4, 3);
  outputView.setUint16(6, 1);
  outputView.setUint32(8, 12);
  output.set(
    bytes.subarray(candidateOffset, candidateOffset + subtableLength),
    12,
  );
  return output;
}

/** 读取并验证 SFNT 表目录，越界或重复表返回空。 */
function readTables(bytes: Uint8Array) {
  if (bytes.length < 12) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const tableCount = view.getUint16(4);
  if (12 + tableCount * 16 > bytes.length) return undefined;
  const tables: SfntTable[] = [];
  const tags = new Set<string>();
  for (let index = 0; index < tableCount; index += 1) {
    const recordOffset = 12 + index * 16;
    const tag = readTag(bytes, recordOffset);
    const offset = view.getUint32(recordOffset + 8);
    const length = view.getUint32(recordOffset + 12);
    if (tags.has(tag) || offset + length > bytes.length) return undefined;
    tags.add(tag);
    tables.push({ tag, offset, length });
  }
  return tables;
}

/**
 * 重建旧式 TTF/OTF 的核心表目录，供浏览器拒绝非必要附加表时降级加载。
 * 原字体可以直接加载时不应调用，避免丢失高级排版与提示信息。
 */
export function createBrowserSafeSfnt(bytes: Uint8Array) {
  const tables = readTables(bytes)
    ?.filter((table) => BROWSER_SAFE_SFNT_TABLES.has(table.tag))
    .sort((left, right) => left.tag.localeCompare(right.tag));
  if (!tables?.length) return undefined;
  const tags = new Set(tables.map((table) => table.tag));
  const hasRequiredMetrics = ['cmap', 'head', 'hhea', 'hmtx', 'maxp'].every(
    (tag) => tags.has(tag),
  );
  const hasGlyphData =
    (tags.has('glyf') && tags.has('loca')) ||
    tags.has('CFF ') ||
    tags.has('CFF2');
  if (!hasRequiredMetrics || !hasGlyphData) return undefined;

  const rebuiltTables = tables.map((table) => {
    const original = bytes.slice(table.offset, table.offset + table.length);
    return {
      tag: table.tag,
      bytes:
        table.tag === 'cmap' ? createBmpCmap(original) ?? original : original,
    };
  });
  const tableCount = rebuiltTables.length;
  const highestPowerOfTwo = 2 ** Math.floor(Math.log2(tableCount));
  const directoryLength = 12 + tableCount * 16;
  let nextTableOffset = alignFour(directoryLength);
  const outputLength = rebuiltTables.reduce(
    (length, table) => length + alignFour(table.bytes.length),
    nextTableOffset,
  );
  const output = new Uint8Array(outputLength);
  output.set(bytes.subarray(0, 4), 0);
  const outputView = new DataView(output.buffer);
  outputView.setUint16(4, tableCount);
  outputView.setUint16(6, highestPowerOfTwo * 16);
  outputView.setUint16(8, Math.log2(highestPowerOfTwo));
  outputView.setUint16(10, tableCount * 16 - highestPowerOfTwo * 16);

  let rebuiltHeadOffset: number | undefined;
  rebuiltTables.forEach((table, index) => {
    const recordOffset = 12 + index * 16;
    const tableBytes = table.bytes;
    if (table.tag === 'head' && tableBytes.length >= 12) {
      new DataView(tableBytes.buffer).setUint32(8, 0);
      rebuiltHeadOffset = nextTableOffset;
    }
    writeTag(output, recordOffset, table.tag);
    outputView.setUint32(recordOffset + 4, calculateChecksum(tableBytes));
    outputView.setUint32(recordOffset + 8, nextTableOffset);
    outputView.setUint32(recordOffset + 12, tableBytes.length);
    output.set(tableBytes, nextTableOffset);
    nextTableOffset += alignFour(tableBytes.length);
  });

  if (rebuiltHeadOffset === undefined) return undefined;
  const adjustment = (SFNT_CHECKSUM_MAGIC - calculateChecksum(output)) >>> 0;
  outputView.setUint32(rebuiltHeadOffset + 8, adjustment);
  return output;
}
