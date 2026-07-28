import { PPT_RECORD } from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import type { PptRecord } from '../types';

/** 从 FontEntityAtom 的 UTF-16LE 定长字段读取字体名称。 */
function readTypeface(record: PptRecord) {
  if (record.data.length < 64) return undefined;
  const value = new TextDecoder('utf-16le')
    .decode(record.data.subarray(0, 64))
    .split('\u0000', 1)[0]
    .trim();
  return value || undefined;
}

/** 递归遍历容器记录，恢复 PPT 字体引用编号与字体名称的映射。 */
export function readPptFonts(
  documentStream: Uint8Array,
  documentRecord: PptRecord,
) {
  const fonts = new Map<number, string>();
  const visit = (record: PptRecord) => {
    if (record.type === PPT_RECORD.FONT_ENTITY_ATOM) {
      const typeface = readTypeface(record);
      if (typeface) fonts.set(record.instance, typeface);
      return;
    }
    if (record.version !== 0x000f || record.length < 8) return;
    for (const child of new PptRecordReader(
      documentStream,
      record.dataOffset,
      record.endOffset,
    ).records()) {
      visit(child);
    }
  };
  visit(documentRecord);
  return fonts;
}
