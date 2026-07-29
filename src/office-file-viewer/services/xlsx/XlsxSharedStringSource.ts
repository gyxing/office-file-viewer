import type { OfficeArchiveReader } from '../../shared/ooxml/OfficeArchiveReader';
import { readOfficeXmlEvents } from '../../shared/ooxml/OfficeXmlEventReader';
import {
  throwIfSpreadsheetAborted,
  waitForSpreadsheetResult,
} from '../spreadsheet/SpreadsheetSource';
import { decodeMojibake } from './parseXlsx';

/** 按当前 Sheet 实际引用读取 sharedStrings 的协议。 */
export interface XlsxSharedStringSource {
  resolveMany(
    indexes: readonly number[],
    signal?: AbortSignal,
  ): Promise<ReadonlyMap<number, string>>;
  dispose(): Promise<void>;
}

/** 创建可重复扫描、只缓存已请求索引的 sharedStrings Source。 */
export function createXlsxSharedStringSource(
  reader: OfficeArchiveReader,
): XlsxSharedStringSource {
  const cache = new Map<number, string>();
  let queue = Promise.resolve();
  let disposed = false;

  const scan = async (requested: Set<number>, signal?: AbortSignal) => {
    if (!requested.size || !reader.has('xl/sharedStrings.xml')) return;
    const stream = await reader.openStream('xl/sharedStrings.xml', signal);
    let index = -1;
    let insideString = false;
    let insideText = false;
    let value = '';
    for await (const event of readOfficeXmlEvents(stream, signal)) {
      throwIfSpreadsheetAborted(signal);
      if (event.type === 'open') {
        if (event.localName === 'si') {
          index += 1;
          insideString = true;
          value = '';
        } else if (insideString && event.localName === 't') {
          insideText = true;
        }
      } else if (event.type === 'text' && insideText) {
        value += event.text;
      } else if (event.type === 'close') {
        if (event.localName === 't') insideText = false;
        if (event.localName === 'si') {
          if (requested.has(index)) {
            cache.set(index, decodeMojibake(value));
            requested.delete(index);
            if (!requested.size) break;
          }
          insideString = false;
          value = '';
        }
      }
    }
  };

  return {
    resolveMany(indexes, signal) {
      if (disposed) {
        return Promise.reject(new Error('XLSX sharedStrings Source 已释放'));
      }
      throwIfSpreadsheetAborted(signal);
      const requested = new Set(
        indexes.filter((index) => Number.isInteger(index) && !cache.has(index)),
      );
      const task = queue.then(
        () => scan(requested, signal),
        () => scan(requested, signal),
      );
      queue = task.catch(() => undefined);
      return waitForSpreadsheetResult(
        task.then(
          () =>
            new Map(
              indexes.map((index) => [index, cache.get(index) ?? ''] as const),
            ),
        ),
        signal,
      );
    },
    async dispose() {
      disposed = true;
      cache.clear();
      await queue;
    },
  };
}
