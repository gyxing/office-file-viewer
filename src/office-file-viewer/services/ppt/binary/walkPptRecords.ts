import type { PptRecord } from '../types';
import { PPT_RECORD } from './constants';
import { PptRecordReader } from './PptRecordReader';

/** 深度遍历 PPT 容器和可识别的可编程二进制标签内容。 */
export function walkPptRecords(
  root: PptRecord,
  visit: (record: PptRecord) => void,
) {
  const walk = (record: PptRecord, depth: number) => {
    visit(record);
    if (
      depth >= 24 ||
      (record.version !== 0x0f &&
        record.type !== PPT_RECORD.BINARY_TAG_DATA_BLOB)
    ) {
      return;
    }
    try {
      for (const child of new PptRecordReader(record.data).records()) {
        walk(child, depth + 1);
      }
    } catch {
      // 未知扩展即使使用容器版本也可能不是记录序列，只忽略当前分支。
    }
  };
  walk(root, 0);
}
