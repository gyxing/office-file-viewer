import { PPT_RECORD } from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import { PptParseError } from '../errors';

/** PPT UserEditAtom 保存的编辑链信息。 */
export type PptUserEdit = {
  /** 在所属数据范围中的偏移位置。 */
  offset: number;
  /** 上一个 UserEditAtom 在主流中的字节偏移。 */
  offsetLastEdit: number;
  /** 持久化目录记录在主流中的字节偏移。 */
  offsetPersistDirectory: number;
  /** PPT 根文档对象的持久化标识。 */
  documentPersistId: number;
  /** 分配后续持久化对象标识时使用的起始值。 */
  persistIdSeed: number;
  /** 加密会话对象的持久化标识；0 表示未加密。 */
  encryptSessionPersistId?: number;
};

/** 从 PowerPoint Document 流的绝对偏移读取一次用户编辑记录。 */
export function readPptUserEdit(
  documentStream: Uint8Array,
  offset: number,
): PptUserEdit {
  const record = new PptRecordReader(
    documentStream,
    offset,
    documentStream.length,
  ).readRecord();
  if (!record || record.type !== PPT_RECORD.USER_EDIT_ATOM) {
    throw new PptParseError(
      'PPT_EDIT_CHAIN_INVALID',
      '编辑链指向的位置不是 UserEditAtom',
      { offset, recordType: record?.type },
    );
  }
  if (record.length !== 28 && record.length !== 32) {
    throw new PptParseError('PPT_EDIT_CHAIN_INVALID', 'UserEditAtom 长度无效', {
      offset,
      recordType: record.type,
    });
  }

  const view = new DataView(
    record.data.buffer,
    record.data.byteOffset,
    record.data.byteLength,
  );
  const offsetLastEdit = view.getUint32(8, true);
  const offsetPersistDirectory = view.getUint32(12, true);
  if (
    offsetLastEdit >= offset ||
    offsetPersistDirectory <= offsetLastEdit ||
    offsetPersistDirectory >= offset
  ) {
    throw new PptParseError(
      'PPT_EDIT_CHAIN_INVALID',
      'UserEditAtom 中的编辑或持久化目录偏移无效',
      { offset, recordType: record.type },
    );
  }

  return {
    offset,
    offsetLastEdit,
    offsetPersistDirectory,
    documentPersistId: view.getUint32(16, true),
    persistIdSeed: view.getUint32(20, true),
    encryptSessionPersistId:
      record.length === 32 ? view.getUint32(28, true) : undefined,
  };
}
