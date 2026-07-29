import type { CfbStreamReader } from '../../../shared/binary/cfb';
import { PPT_RECORD } from '../binary/constants';
import { PptParseError } from '../errors';
import { readPptRecordAtOffset } from '../readPptPersistObject';
import type { PptEditChain, PptParseContext } from '../types';
import { readPptCurrentUser } from './readCurrentUser';

function parseUserEdit(
  bytes: Uint8Array,
  sourceOffset: number,
): {
  offsetLastEdit: number;
  offsetPersistDirectory: number;
  documentPersistId: number;
  persistIdSeed: number;
  encryptSessionPersistId?: number;
} {
  const payload = bytes.subarray(8);
  if (payload.length !== 28 && payload.length !== 32) {
    throw new PptParseError('PPT_EDIT_CHAIN_INVALID', 'UserEditAtom 长度无效', {
      offset: sourceOffset,
      recordType: PPT_RECORD.USER_EDIT_ATOM,
    });
  }
  const view = new DataView(
    payload.buffer,
    payload.byteOffset,
    payload.byteLength,
  );
  const offsetLastEdit = view.getUint32(8, true);
  const offsetPersistDirectory = view.getUint32(12, true);
  if (
    offsetLastEdit >= sourceOffset ||
    offsetPersistDirectory <= offsetLastEdit ||
    offsetPersistDirectory >= sourceOffset
  ) {
    throw new PptParseError(
      'PPT_EDIT_CHAIN_INVALID',
      'UserEditAtom 中的编辑或持久化目录偏移无效',
      { offset: sourceOffset, recordType: PPT_RECORD.USER_EDIT_ATOM },
    );
  }
  return {
    offsetLastEdit,
    offsetPersistDirectory,
    documentPersistId: view.getUint32(16, true),
    persistIdSeed: view.getUint32(20, true),
    encryptSessionPersistId:
      payload.length === 32 ? view.getUint32(28, true) : undefined,
  };
}

function parsePersistDirectory(
  bytes: Uint8Array,
  sourceOffset: number,
  lastEditOffset: number,
  streamSize: number,
) {
  const payload = bytes.subarray(8);
  const view = new DataView(
    payload.buffer,
    payload.byteOffset,
    payload.byteLength,
  );
  const result = new Map<number, number>();
  let cursor = 0;
  while (cursor < payload.length) {
    if (payload.length - cursor < 4) {
      throw new PptParseError(
        'PPT_PERSIST_DIRECTORY_INVALID',
        'PersistDirectoryEntry 头不完整',
        { offset: sourceOffset + 8 + cursor },
      );
    }
    const descriptor = view.getUint32(cursor, true);
    cursor += 4;
    const persistId = descriptor & 0x000fffff;
    const count = descriptor >>> 20;
    if (!count || count > (payload.length - cursor) / 4) {
      throw new PptParseError(
        'PPT_PERSIST_DIRECTORY_INVALID',
        'PersistDirectoryEntry 数量超出记录边界',
        { offset: sourceOffset + 8 + cursor - 4 },
      );
    }
    for (let index = 0; index < count; index += 1) {
      const persistOffset = view.getUint32(cursor, true);
      cursor += 4;
      if (
        persistOffset < lastEditOffset ||
        persistOffset >= sourceOffset ||
        persistOffset >= streamSize
      ) {
        throw new PptParseError(
          'PPT_PERSIST_DIRECTORY_INVALID',
          'Persist 对象偏移超出当前编辑范围',
          { offset: sourceOffset + 8 + cursor - 4 },
        );
      }
      result.set(persistId + index, persistOffset);
    }
  }
  return result;
}

/** 只按编辑链偏移读取 UserEdit 和 PersistDirectory，不物化 PPT 主流。 */
export async function buildPptEditChainFromStream(
  documentStream: CfbStreamReader,
  currentUserStream: Uint8Array,
  context: PptParseContext,
  signal?: AbortSignal,
): Promise<PptEditChain> {
  const currentUser = readPptCurrentUser(currentUserStream, context);
  const visited = new Set<number>();
  const editOffsets: number[] = [];
  const persistOffsets = new Map<number, number>();
  let editOffset = currentUser.offsetToCurrentEdit;
  let documentPersistId = 0;
  let persistIdSeed = 0;

  while (editOffset) {
    if (visited.has(editOffset)) {
      throw new PptParseError(
        'PPT_EDIT_CHAIN_CYCLE',
        'PowerPoint 增量保存链存在循环',
        { offset: editOffset },
      );
    }
    visited.add(editOffset);
    editOffsets.push(editOffset);
    const editRecord = await readPptRecordAtOffset(
      documentStream,
      editOffset,
      signal,
    );
    if (editRecord.type !== PPT_RECORD.USER_EDIT_ATOM) {
      throw new PptParseError(
        'PPT_EDIT_CHAIN_INVALID',
        '编辑链指向的位置不是 UserEditAtom',
        { offset: editOffset, recordType: editRecord.type },
      );
    }
    const edit = parseUserEdit(editRecord.bytes, editOffset);
    if (edit.encryptSessionPersistId) {
      throw new PptParseError('PPT_ENCRYPTED', '暂不支持加密的 PPT 文件', {
        offset: editOffset,
        recordType: PPT_RECORD.USER_EDIT_ATOM,
      });
    }
    if (!documentPersistId) documentPersistId = edit.documentPersistId;
    persistIdSeed = Math.max(persistIdSeed, edit.persistIdSeed);
    const directoryRecord = await readPptRecordAtOffset(
      documentStream,
      edit.offsetPersistDirectory,
      signal,
    );
    if (
      directoryRecord.type !== PPT_RECORD.PERSIST_PTR_FULL_BLOCK &&
      directoryRecord.type !== PPT_RECORD.PERSIST_PTR_INCREMENTAL_BLOCK
    ) {
      throw new PptParseError(
        'PPT_PERSIST_DIRECTORY_INVALID',
        '编辑链指向的位置不是 PersistDirectoryAtom',
        {
          offset: edit.offsetPersistDirectory,
          recordType: directoryRecord.type,
        },
      );
    }
    parsePersistDirectory(
      directoryRecord.bytes,
      edit.offsetPersistDirectory,
      edit.offsetLastEdit,
      documentStream.entry.streamSize,
    ).forEach((offset, persistId) => {
      if (!persistOffsets.has(persistId)) persistOffsets.set(persistId, offset);
    });
    editOffset = edit.offsetLastEdit;
    await context.yieldIfNeeded();
  }

  const documentOffset = persistOffsets.get(documentPersistId);
  if (documentOffset === undefined) {
    throw new PptParseError(
      'PPT_DOCUMENT_MISSING',
      '持久化目录中缺少根文档对象',
    );
  }
  const documentRecord = await readPptRecordAtOffset(
    documentStream,
    documentOffset,
    signal,
  );
  if (documentRecord.type !== PPT_RECORD.DOCUMENT) {
    throw new PptParseError(
      'PPT_DOCUMENT_MISSING',
      '根持久化对象不是 DocumentContainer',
      { offset: documentOffset, recordType: documentRecord.type },
    );
  }

  return {
    documentPersistId,
    persistIdSeed,
    persistOffsets,
    editOffsets,
  };
}
