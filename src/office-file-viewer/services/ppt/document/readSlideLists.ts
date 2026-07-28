import { PPT_RECORD } from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import type { PptParseContext, PptRecord, PptSlideDescriptor } from '../types';

/** 描述 PptMasterDescriptor 在 PPT 二进制解析中的数据结构。 */
export type PptMasterDescriptor = {
  /** PptMasterDescriptor 在源文件记录中的数字标识。 */
  persistId: number;
  /** PptMasterDescriptor 在源文件记录中的数字标识。 */
  masterId: number;
};

/** 备注列表中的持久化引用；实际幻灯片关联必须读取 NotesAtom.slideIdRef。 */
export type PptNotesDescriptor = {
  /** NotesContainer 在持久化目录中的引用。 */
  persistId: number;
  /** 备注页自身标识，仅用于诊断，不能作为幻灯片顺序。 */
  notesId: number;
};

/** 读取 `readSlidePersistFields` 所需的源数据，供PPT 二进制解析使用。 */
function readSlidePersistFields(record: PptRecord) {
  if (record.length < 16) return undefined;
  const view = new DataView(
    record.data.buffer,
    record.data.byteOffset,
    record.data.byteLength,
  );
  return {
    persistId: view.getUint32(0, true),
    objectId: view.getUint32(12, true),
  };
}

/** 从指定实例的 SlideListWithTextContainer 恢复页面或母版引用顺序。 */
export function readPptSlideLists(
  documentStream: Uint8Array,
  documentRecord: PptRecord,
  context: PptParseContext,
) {
  const slides: PptSlideDescriptor[] = [];
  const masters: PptMasterDescriptor[] = [];
  const notes: PptNotesDescriptor[] = [];
  const reader = new PptRecordReader(
    documentStream,
    documentRecord.dataOffset,
    documentRecord.endOffset,
  );

  for (const child of reader.records()) {
    if (child.type !== PPT_RECORD.SLIDE_LIST_WITH_TEXT) continue;
    if (child.instance !== 0 && child.instance !== 1 && child.instance !== 2)
      continue;
    const listReader = new PptRecordReader(
      documentStream,
      child.dataOffset,
      child.endOffset,
    );
    for (const item of listReader.records()) {
      if (item.type !== PPT_RECORD.SLIDE_PERSIST_ATOM) continue;
      const fields = readSlidePersistFields(item);
      if (!fields?.persistId) {
        context.warnings.push({
          code: 'PPT_SLIDE_PERSIST_INVALID',
          message: '幻灯片、母版或备注引用记录长度无效',
          offset: item.offset,
        });
        continue;
      }
      if (child.instance === 0) {
        slides.push({
          persistId: fields.persistId,
          slideId: fields.objectId,
          index: slides.length + 1,
        });
      } else if (child.instance === 1) {
        masters.push({
          persistId: fields.persistId,
          masterId: fields.objectId,
        });
      } else {
        notes.push({
          persistId: fields.persistId,
          notesId: fields.objectId,
        });
      }
    }
  }

  return { slides, masters, notes };
}
