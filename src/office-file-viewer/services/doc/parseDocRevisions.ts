import type { DocCharacterRevision, DocFib } from './docParseTypes';

/** 读取普通 SPRM 操作数长度；字符修订不使用表格专属变体。 */
function readOperandSize(sprm: number, bytes: Uint8Array, offset: number) {
  const sizeCode = (sprm >> 13) & 0x7;
  if (sizeCode === 0 || sizeCode === 1) return 1;
  if (sizeCode === 2 || sizeCode === 4 || sizeCode === 5) return 2;
  if (sizeCode === 3) return 4;
  if (sizeCode === 6) return 1 + (bytes[offset] ?? 0);
  if (sizeCode === 7) return 3;
  return 0;
}

/** 将 Word DTTM 位字段转换为稳定的本地 ISO 日期文本。 */
function decodeDocRevisionDate(value: number) {
  const minute = value & 0x3f;
  const hour = (value >>> 6) & 0x1f;
  const day = (value >>> 11) & 0x1f;
  const month = (value >>> 16) & 0x0f;
  const year = 1900 + ((value >>> 20) & 0x1ff);
  if (
    year < 1900 ||
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > 31 ||
    hour > 23 ||
    minute > 59
  ) {
    return undefined;
  }
  const pad = (part: number) => `${part}`.padStart(2, '0');
  return `${year}-${pad(month)}-${pad(day)}T${pad(hour)}:${pad(minute)}`;
}

/** 兼容 Unicode 和 ANSI STTB 的修订作者名称表。 */
export function readDocRevisionAuthors(tableStream: Uint8Array, fib: DocFib) {
  const start = fib.fcSttbfRMark;
  const end = start + fib.lcbSttbfRMark;
  if (!start || end > tableStream.length || end - start < 4) return [];
  const data = tableStream.slice(start, end);
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const extended = view.getUint16(0, true) === 0xffff;
  const count = view.getUint16(extended ? 2 : 0, true);
  const extraSize = view.getUint16(extended ? 4 : 2, true);
  let offset = extended ? 6 : 4;
  const authors: string[] = [];
  for (let index = 0; index < count; index += 1) {
    const lengthFieldSize = extended ? 2 : 1;
    if (offset + lengthFieldSize > data.length) break;
    const charCount = extended
      ? view.getUint16(offset, true)
      : data[offset] ?? 0;
    offset += lengthFieldSize;
    const byteLength = charCount * (extended ? 2 : 1);
    if (offset + byteLength + extraSize > data.length) break;
    const bytes = data.slice(offset, offset + byteLength);
    authors.push(
      new TextDecoder(extended ? 'utf-16le' : 'windows-1252')
        .decode(bytes)
        .replace(/\u0000/g, '')
        .trim(),
    );
    offset += byteLength + extraSize;
  }
  return authors;
}

/** 从单个 CHPX grpprl 读取插入、删除、作者和日期修订属性。 */
export function parseDocCharacterRevisions(bytes: Uint8Array) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = 0;
  let deleted = false;
  let inserted = false;
  let authorIndex: number | undefined;
  let createdAt: string | undefined;
  while (offset + 2 <= bytes.length) {
    const sprm = view.getUint16(offset, true);
    offset += 2;
    const operandSize = readOperandSize(sprm, bytes, offset);
    if (!operandSize || offset + operandSize > bytes.length) break;
    if (sprm === 0x0800) deleted = Boolean((bytes[offset] ?? 0) & 1);
    if (sprm === 0x0801) inserted = Boolean((bytes[offset] ?? 0) & 1);
    if (sprm === 0x4804 && operandSize >= 2) {
      authorIndex = view.getUint16(offset, true);
    }
    if (sprm === 0x6805 && operandSize >= 4) {
      createdAt = decodeDocRevisionDate(view.getUint32(offset, true));
    }
    offset += operandSize;
  }
  const revisions: DocCharacterRevision[] = [];
  if (deleted) revisions.push({ kind: 'delete', authorIndex, createdAt });
  if (inserted) revisions.push({ kind: 'insert', authorIndex, createdAt });
  return revisions;
}

/** 判断 grpprl 是否包含尚未还原原始格式值的属性级修订。 */
export function hasDocPropertyRevision(bytes: Uint8Array) {
  const propertyRevisionSprms = new Set([
    0xca57, // Word 97 字符属性修订
    0xca89, // 新版字符属性修订
    0xc66f, // 段落属性修订
    0xd667, // 表格属性修订
    0xd227, // Word 97 节属性修订
    0xd243, // 新版节属性修订
  ]);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = 0;
  while (offset + 2 <= bytes.length) {
    const sprm = view.getUint16(offset, true);
    offset += 2;
    const operandSize = readOperandSize(sprm, bytes, offset);
    if (!operandSize || offset + operandSize > bytes.length) break;
    const payloadOffset = ((sprm >> 13) & 0x7) === 6 ? offset + 1 : offset;
    if (
      propertyRevisionSprms.has(sprm) &&
      payloadOffset < offset + operandSize &&
      bytes[payloadOffset] !== 0
    ) {
      return true;
    }
    offset += operandSize;
  }
  return false;
}
