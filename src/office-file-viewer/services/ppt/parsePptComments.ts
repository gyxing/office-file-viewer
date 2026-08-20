import type {
  PresentationAnnotation,
  SlideElement,
} from '../presentation/types';
import { PPT_RECORD } from './binary/constants';
import { PptRecordReader } from './binary/PptRecordReader';
import { readPptUnicodeString } from './binary/readStrings';
import { walkPptRecords } from './binary/walkPptRecords';
import type { PptRecord } from './types';

/** PowerPoint 主坐标固定为 576 dpi，统一换算到浏览器 96 dpi。 */
const MASTER_UNIT_TO_PX = 96 / 576;

function readCreatedAt(bytes: Uint8Array) {
  if (bytes.length < 20) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const year = view.getUint16(4, true);
  const month = view.getUint16(6, true);
  const day = view.getUint16(10, true);
  const hour = view.getUint16(12, true);
  const minute = view.getUint16(14, true);
  const second = view.getUint16(16, true);
  const millisecond = view.getUint16(18, true);
  if (!year || month < 1 || month > 12 || day < 1 || day > 31) {
    return undefined;
  }
  return new Date(
    Date.UTC(year, month - 1, day, hour, minute, second, millisecond),
  ).toISOString();
}

function findNearestElement(
  elements: readonly SlideElement[],
  x: number,
  y: number,
) {
  return elements
    .map((element) => {
      const within =
        x >= element.x &&
        x <= element.x + element.width &&
        y >= element.y &&
        y <= element.y + element.height;
      const dx = x - (element.x + element.width / 2);
      const dy = y - (element.y + element.height / 2);
      return { id: element.id, distance: within ? -1 : Math.hypot(dx, dy) };
    })
    .sort((left, right) => left.distance - right.distance)[0]?.id;
}

function parseCommentContainer(
  record: PptRecord,
  slideId: string,
  slideIndex: number,
  elements: readonly SlideElement[],
): PresentationAnnotation | undefined {
  const children = Array.from(new PptRecordReader(record.data).records());
  const strings = children
    .filter((child) => child.type === PPT_RECORD.C_STRING)
    .map((child) => readPptUnicodeString(child.data));
  const atom = children.find(
    (child) => child.type === PPT_RECORD.COMMENT_10_ATOM,
  );
  const text = strings[1]?.trim();
  if (!atom || atom.data.length < 28 || !text) return undefined;
  const view = new DataView(
    atom.data.buffer,
    atom.data.byteOffset,
    atom.data.byteLength,
  );
  const index = view.getInt32(0, true);
  const x = view.getInt32(20, true) * MASTER_UNIT_TO_PX;
  const y = view.getInt32(24, true) * MASTER_UNIT_TO_PX;
  return {
    id: `ppt-comment-${slideIndex + 1}-${index}`,
    author: strings[0]?.trim() || undefined,
    createdAt: readCreatedAt(atom.data),
    text,
    x,
    y,
    target: {
      kind: 'presentation-element',
      slideId,
      slideIndex,
      elementId: findNearestElement(elements, x, y),
    },
  };
}

/** 从 PPT10 可编程标签中恢复当前幻灯片批注。 */
export function parsePptComments(
  slideRecord: PptRecord,
  slideId: string,
  slideIndex: number,
  elements: readonly SlideElement[],
) {
  const annotations: PresentationAnnotation[] = [];
  walkPptRecords(slideRecord, (record) => {
    if (record.type !== PPT_RECORD.COMMENT_10) return;
    try {
      const annotation = parseCommentContainer(
        record,
        slideId,
        slideIndex,
        elements,
      );
      if (annotation) annotations.push(annotation);
    } catch {
      // 损坏的单条批注不应中断幻灯片正文和其他批注。
    }
  });
  return annotations;
}
