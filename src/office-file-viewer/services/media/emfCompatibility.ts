import type { EmfPatternBrush, EmfRecord } from './emfCompatibilityBinary';
import {
  convertAlphaBlend,
  convertPatternAndBitBlt,
  convertPolyPolyline16,
  createPatternBrushPlaceholder,
  EMR_ALPHABLEND,
  EMR_BITBLT,
  EMR_CREATEDIBPATTERNBRUSHPT,
  EMR_POLYPOLYLINE16,
  hasEmfCompatibilityRecords,
  readPatternBrush,
  rebuildEmf,
  scanEmfRecords,
} from './emfCompatibilityBinary';

/** 保存当前设备上下文。 */
const EMR_SAVEDC = 33;
/** 恢复先前保存的设备上下文。 */
const EMR_RESTOREDC = 34;
/** 设置图案画刷的平铺原点。 */
const EMR_SETBRUSHORGEX = 13;
/** 在设备上下文中选择图形对象。 */
const EMR_SELECTOBJECT = 37;
/** 创建普通画笔。 */
const EMR_CREATEPEN = 38;
/** 创建普通画刷。 */
const EMR_CREATEBRUSHINDIRECT = 39;
/** 从对象表删除图形对象。 */
const EMR_DELETEOBJECT = 40;
/** 创建逻辑字体。 */
const EMR_EXTCREATEFONTINDIRECTW = 82;
/** 创建单色图案画刷。 */
const EMR_CREATEMONOBRUSH = 93;
/** 创建扩展画笔。 */
const EMR_EXTCREATEPEN = 95;
/** EMF 预定义对象句柄的起始值。 */
const STOCK_OBJECT_BASE = 0x80000000;

/** 会使用当前画刷填充内容、但兼容层尚不能展开图案的记录类型。 */
const BRUSH_DRAW_RECORDS = new Set([
  3, 8, 42, 43, 44, 46, 47, 53, 62, 63, 71, 72, 74, 77, 78, 86, 91,
]);

/** 单个图案画刷的兼容性分析结果。 */
type PatternCandidate = {
  /** 已解码的黑白图案。 */
  brush: EmfPatternBrush;
  /** 当前画刷生命周期内的用法是否都能等价转换。 */
  safe: boolean;
  /** 以原记录偏移为键保存已展开的蒙版记录。 */
  replacements: Map<number, Uint8Array>;
};

/** 对象表中与画刷选择有关的对象类别。 */
type TrackedObject =
  | { kind: 'brush' }
  | { kind: 'pen' | 'font' }
  | { kind: 'pattern'; candidate: PatternCandidate };

/** 保存/恢复设备上下文时需要同步的兼容状态。 */
type CompatibilitySnapshot = {
  /** 保存时选中的图案画刷；undefined 表示普通画刷。 */
  selectedPattern?: PatternCandidate;
  /** 保存时的画刷原点横坐标。 */
  brushOriginX: number;
  /** 保存时的画刷原点纵坐标。 */
  brushOriginY: number;
};

function createView(bytes: Uint8Array) {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

function stockObjectKind(handle: number): TrackedObject['kind'] | undefined {
  if (handle < STOCK_OBJECT_BASE) return undefined;
  const index = handle - STOCK_OBJECT_BASE;
  if (index >= 0 && index <= 5) return 'brush';
  if (index >= 6 && index <= 8) return 'pen';
  if (index >= 10 && index <= 17) return 'font';
  return undefined;
}

function markAmbiguousSelectionUnsafe(
  selectedPattern: PatternCandidate | undefined,
) {
  if (selectedPattern) selectedPattern.safe = false;
  return undefined;
}

function restoreSnapshot(snapshots: CompatibilitySnapshot[], relative: number) {
  const index = relative < 0 ? snapshots.length + relative : relative - 1;
  if (index < 0 || index >= snapshots.length) return undefined;
  const snapshot = snapshots[index];
  snapshots.splice(index);
  return snapshot;
}

function registerObject(
  objects: Map<number, TrackedObject>,
  handle: number,
  object: TrackedObject,
) {
  objects.set(handle, object);
}

function trackCreatedObject(
  bytes: Uint8Array,
  record: EmfRecord,
  objects: Map<number, TrackedObject>,
  candidates: PatternCandidate[],
) {
  if (record.size < 12) return;
  const handle = createView(bytes).getUint32(record.offset + 8, true);
  if (record.type === EMR_CREATEDIBPATTERNBRUSHPT) {
    const brush = readPatternBrush(bytes, record);
    if (!brush) {
      registerObject(objects, handle, { kind: 'brush' });
      return;
    }
    const candidate: PatternCandidate = {
      brush,
      safe: true,
      replacements: new Map(),
    };
    candidates.push(candidate);
    registerObject(objects, handle, { kind: 'pattern', candidate });
    return;
  }
  if (
    record.type === EMR_CREATEBRUSHINDIRECT ||
    record.type === EMR_CREATEMONOBRUSH
  ) {
    registerObject(objects, handle, { kind: 'brush' });
    return;
  }
  if (record.type === EMR_EXTCREATEFONTINDIRECTW) {
    registerObject(objects, handle, { kind: 'font' });
    return;
  }
  registerObject(objects, handle, { kind: 'pen' });
}

function isObjectCreationRecord(type: number) {
  return (
    type === EMR_CREATEPEN ||
    type === EMR_CREATEBRUSHINDIRECT ||
    type === EMR_EXTCREATEFONTINDIRECTW ||
    type === EMR_CREATEMONOBRUSH ||
    type === EMR_CREATEDIBPATTERNBRUSHPT ||
    type === EMR_EXTCREATEPEN
  );
}

function collectPatternReplacements(bytes: Uint8Array, records: EmfRecord[]) {
  const view = createView(bytes);
  const objects = new Map<number, TrackedObject>();
  const candidates: PatternCandidate[] = [];
  const snapshots: CompatibilitySnapshot[] = [];
  let selectedPattern: PatternCandidate | undefined;
  let brushOriginX = 0;
  let brushOriginY = 0;

  for (const record of records) {
    if (isObjectCreationRecord(record.type)) {
      trackCreatedObject(bytes, record, objects, candidates);
      continue;
    }
    if (record.type === EMR_SETBRUSHORGEX && record.size >= 16) {
      brushOriginX = view.getInt32(record.offset + 8, true);
      brushOriginY = view.getInt32(record.offset + 12, true);
      continue;
    }
    if (record.type === EMR_SAVEDC) {
      snapshots.push({ selectedPattern, brushOriginX, brushOriginY });
      continue;
    }
    if (record.type === EMR_RESTOREDC && record.size >= 12) {
      const restored = restoreSnapshot(
        snapshots,
        view.getInt32(record.offset + 8, true),
      );
      if (!restored) {
        selectedPattern = markAmbiguousSelectionUnsafe(selectedPattern);
        continue;
      }
      selectedPattern = restored.selectedPattern;
      brushOriginX = restored.brushOriginX;
      brushOriginY = restored.brushOriginY;
      continue;
    }
    if (record.type === EMR_SELECTOBJECT && record.size >= 12) {
      const handle = view.getUint32(record.offset + 8, true);
      const kind = stockObjectKind(handle);
      if (kind === 'brush') {
        selectedPattern = undefined;
        continue;
      }
      if (kind === 'pen' || kind === 'font') continue;
      const object = objects.get(handle);
      if (object?.kind === 'pattern') {
        selectedPattern = object.candidate;
      } else if (object?.kind === 'brush') {
        selectedPattern = undefined;
      } else if (!object) {
        selectedPattern = markAmbiguousSelectionUnsafe(selectedPattern);
      }
      continue;
    }
    if (record.type === EMR_DELETEOBJECT && record.size >= 12) {
      const handle = view.getUint32(record.offset + 8, true);
      const object = objects.get(handle);
      if (object?.kind === 'pattern' && selectedPattern === object.candidate) {
        selectedPattern = undefined;
      }
      objects.delete(handle);
      continue;
    }
    if (record.type === EMR_BITBLT && selectedPattern) {
      const replacement = convertPatternAndBitBlt(bytes, record, {
        brush: selectedPattern.brush,
        originX: brushOriginX,
        originY: brushOriginY,
      });
      if (replacement) {
        selectedPattern.replacements.set(record.offset, replacement);
      } else {
        selectedPattern.safe = false;
      }
      continue;
    }
    if (BRUSH_DRAW_RECORDS.has(record.type) && selectedPattern) {
      selectedPattern.safe = false;
    }
  }

  const replacements = new Map<number, Uint8Array>();
  for (const candidate of candidates) {
    if (!candidate.safe) continue;
    replacements.set(
      candidate.brush.creationOffset,
      createPatternBrushPlaceholder(candidate.brush.handle),
    );
    for (const [offset, replacement] of candidate.replacements) {
      replacements.set(offset, replacement);
    }
  }
  return replacements;
}

/**
 * 将 emf-converter 暂不支持但可等价表达的标准 EMF 记录预转换。
 * 未命中兼容记录时复用原字节，避免普通图片产生额外分配。
 */
export function normalizeEmfForConverter(bytes: Uint8Array) {
  try {
    if (!hasEmfCompatibilityRecords(bytes)) return bytes;
    const records = scanEmfRecords(bytes);
    if (!records) return bytes;
    const replacements = collectPatternReplacements(bytes, records);
    for (const record of records) {
      if (record.type === EMR_POLYPOLYLINE16) {
        const replacement = convertPolyPolyline16(bytes, record);
        if (replacement) replacements.set(record.offset, replacement);
      } else if (record.type === EMR_ALPHABLEND) {
        const replacement = convertAlphaBlend(bytes, record);
        if (replacement) replacements.set(record.offset, replacement);
      }
    }
    return rebuildEmf(bytes, records, replacements);
  } catch {
    // 单条非标准记录不能影响整份 Office 文档，保留原数据交给现有转换器。
    return bytes;
  }
}
