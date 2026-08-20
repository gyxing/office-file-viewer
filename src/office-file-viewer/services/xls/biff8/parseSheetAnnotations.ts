import type { SpreadsheetAnnotation } from '../../spreadsheet/semantics/types';
import type { SpreadsheetWarning } from '../../spreadsheet/types';
import type { Biff8RecordSequence, Biff8SheetDescriptor } from '../types';
import { BIFF8_RECORD } from './constants';

/** 读取小端 16 位整数，越界时返回 0。 */
function readUint16(bytes: Uint8Array, offset: number) {
  if (offset + 2 > bytes.length) return 0;
  return new DataView(
    bytes.buffer,
    bytes.byteOffset,
    bytes.byteLength,
  ).getUint16(offset, true);
}

/** 解码 NOTE 记录尾部的 BIFF8 Unicode 作者。 */
function readNoteAuthor(bytes: Uint8Array) {
  if (bytes.length < 11) return undefined;
  const count = readUint16(bytes, 8);
  const flags = bytes[10] ?? 0;
  const highByte = Boolean(flags & 0x01);
  const start = 11;
  const byteLength = count * (highByte ? 2 : 1);
  if (!count || start + byteLength > bytes.length) return undefined;
  return new TextDecoder(highByte ? 'utf-16le' : 'windows-1252')
    .decode(bytes.slice(start, start + byteLength))
    .replace(/\u0000+$/g, '');
}

/** 从 OBJ 的 FtCmo 子记录读取对象标识。 */
function readObjectId(bytes: Uint8Array) {
  return bytes.length >= 8 && readUint16(bytes, 0) === 0x0015
    ? readUint16(bytes, 6)
    : undefined;
}

/** 从 TXO 与连续 CONTINUE 记录读取批注正文。 */
function readTxoText(sequence: Biff8RecordSequence) {
  const header = sequence.chunks[0];
  const characterCount = readUint16(header, 10);
  if (!characterCount) return '';
  let remaining = characterCount;
  let text = '';
  for (const chunk of sequence.chunks.slice(1)) {
    if (!remaining || !chunk.length) break;
    const highByte = Boolean((chunk[0] ?? 0) & 0x01);
    const available = Math.floor((chunk.length - 1) / (highByte ? 2 : 1));
    const count = Math.min(remaining, available);
    if (!count) continue;
    text += new TextDecoder(highByte ? 'utf-16le' : 'windows-1252').decode(
      chunk.slice(1, 1 + count * (highByte ? 2 : 1)),
    );
    remaining -= count;
  }
  return text.replace(/\u0000+$/g, '');
}

/** 关联 OBJ、TXO 和 NOTE，生成按单元格定位的 BIFF8 批注。 */
export function parseBiff8SheetAnnotations(
  descriptor: Biff8SheetDescriptor,
  records: readonly Biff8RecordSequence[],
  warnings: SpreadsheetWarning[],
) {
  const textByObjectId = new Map<number, string>();
  let activeObjectId: number | undefined;
  records.forEach((record) => {
    if (record.recordId === BIFF8_RECORD.OBJ) {
      activeObjectId = readObjectId(record.chunks[0]);
      return;
    }
    if (record.recordId === BIFF8_RECORD.TXO && activeObjectId !== undefined) {
      textByObjectId.set(activeObjectId, readTxoText(record));
    }
  });
  const annotations: SpreadsheetAnnotation[] = [];
  records.forEach((record, index) => {
    if (record.recordId !== BIFF8_RECORD.NOTE) return;
    const bytes = record.chunks[0];
    if (bytes.length < 8) {
      warnings.push({
        code: 'INVALID_NOTE',
        message: 'NOTE 记录长度无效，已跳过当前单元格批注',
        sheetName: descriptor.name,
        offset: record.offset,
      });
      return;
    }
    const row = readUint16(bytes, 0) + 1;
    const column = readUint16(bytes, 2) + 1;
    const objectId = readUint16(bytes, 6);
    const ref = `${columnLabel(column)}${row}`;
    annotations.push({
      id: `${descriptor.id}:note:${objectId || index + 1}`,
      ref,
      row,
      column,
      author: readNoteAuthor(bytes),
      text: textByObjectId.get(objectId) ?? '',
    });
  });
  return annotations;
}

/** 将一基列号转换为 A1 列标签。 */
function columnLabel(index: number) {
  let value = Math.max(1, Math.trunc(index));
  let label = '';
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return label;
}
