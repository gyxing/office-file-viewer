import type { DocTextStyle } from './types';

/** 二进制 DOC 的单级列表格式。 */
type DocListLevel = {
  /** 当前范围的起始位置。 */
  start: number;
  /** 源文件的格式标识。 */
  format: number;
  /** 文本内容。 */
  text: string;
  /** 列表编号与正文之间使用的后缀类型。 */
  suffix: 'tab' | 'space' | 'nothing';
  /** 列表级别声明的悬挂缩进等段落布局。 */
  style?: DocTextStyle;
};

/** 二进制 DOC 列表定义、覆盖实例与当前计数状态。 */
export type DocNumberingCatalog = {
  /** 按列表标识索引的多级编号定义。 */
  lists: Map<number, DocListLevel[]>;
  /** 各列表实例引用的列表定义标识。 */
  instanceListIds: number[];
  /** 按列表实例保存的各级当前计数器。 */
  counters: Map<number, number[]>;
};

function readUint16(bytes: Uint8Array, offset: number) {
  if (offset < 0 || offset + 2 > bytes.length) return undefined;
  return new DataView(
    bytes.buffer,
    bytes.byteOffset,
    bytes.byteLength,
  ).getUint16(offset, true);
}

function readInt16(bytes: Uint8Array, offset: number) {
  if (offset < 0 || offset + 2 > bytes.length) return undefined;
  return new DataView(
    bytes.buffer,
    bytes.byteOffset,
    bytes.byteLength,
  ).getInt16(offset, true);
}

function readUint32(bytes: Uint8Array, offset: number) {
  if (offset < 0 || offset + 4 > bytes.length) return undefined;
  return new DataView(
    bytes.buffer,
    bytes.byteOffset,
    bytes.byteLength,
  ).getUint32(offset, true);
}

function readInt32(bytes: Uint8Array, offset: number) {
  if (offset < 0 || offset + 4 > bytes.length) return undefined;
  return new DataView(
    bytes.buffer,
    bytes.byteOffset,
    bytes.byteLength,
  ).getInt32(offset, true);
}

function twipToPx(value: number) {
  return (value / 1440) * 96;
}

/** 跳过列表级别 PAPX 中当前 SPRM 的操作数。 */
function listSprmOperandSize(sprm: number, bytes: Uint8Array, offset: number) {
  const sizeCode = (sprm >> 13) & 0x7;
  if (sizeCode === 0 || sizeCode === 1) return 1;
  if (sizeCode === 2 || sizeCode === 4 || sizeCode === 5) return 2;
  if (sizeCode === 3) return 4;
  if (sizeCode === 6) return 1 + (bytes[offset] ?? 0);
  if (sizeCode === 7) return 3;
  return 0;
}

/** 读取列表级别保存的悬挂缩进，避免 HTML 默认列表缩进改变源换行。 */
function readListParagraphStyle(bytes: Uint8Array) {
  const style: DocTextStyle = {};
  let offset = 0;
  while (offset + 2 <= bytes.length) {
    const sprm = readUint16(bytes, offset);
    if (sprm === undefined) break;
    offset += 2;
    const operandSize = listSprmOperandSize(sprm, bytes, offset);
    if (!operandSize || offset + operandSize > bytes.length) break;
    const value = readInt16(bytes, offset);
    if (value !== undefined && (sprm === 0x840f || sprm === 0x845e)) {
      style.indentLeft = twipToPx(value);
    } else if (value !== undefined && (sprm === 0x8411 || sprm === 0x8460)) {
      style.firstLineIndent = twipToPx(value);
    }
    offset += operandSize;
  }
  return Object.keys(style).length ? style : undefined;
}

function readListLevel(
  tableStream: Uint8Array,
  offset: number,
): { level: DocListLevel; nextOffset: number } | undefined {
  if (offset < 0 || offset + 30 > tableStream.length) return undefined;
  const start = readInt32(tableStream, offset) ?? 1;
  const format = tableStream[offset + 4] ?? 0;
  const follow = tableStream[offset + 15] ?? 0;
  const characterPropertyLength = tableStream[offset + 24] ?? 0;
  const paragraphPropertyLength = tableStream[offset + 25] ?? 0;
  const paragraphProperties = tableStream.slice(
    offset + 28,
    offset + 28 + paragraphPropertyLength,
  );
  const textLengthOffset =
    offset + 28 + paragraphPropertyLength + characterPropertyLength;
  const textLength = readUint16(tableStream, textLengthOffset);
  if (
    textLength === undefined ||
    textLengthOffset + 2 + textLength * 2 > tableStream.length
  ) {
    return undefined;
  }
  const textBytes = tableStream.slice(
    textLengthOffset + 2,
    textLengthOffset + 2 + textLength * 2,
  );
  const rawText = new TextDecoder('utf-16le').decode(textBytes);
  const text = Array.from(rawText)
    .map((character) => {
      const code = character.charCodeAt(0);
      return code <= 8 ? `%${code + 1}` : character;
    })
    .join('');
  return {
    level: {
      start: Math.max(0, start),
      format,
      text,
      suffix: follow === 1 ? 'space' : follow === 2 ? 'nothing' : 'tab',
      style: readListParagraphStyle(paragraphProperties),
    },
    nextOffset: textLengthOffset + 2 + textLength * 2,
  };
}

/** 从 FIB 指向的 PlfLst 与 PlfLfo 读取列表格式。 */
export function readDocNumberingCatalog(
  tableStream: Uint8Array,
  fcPlfLst: number,
  lcbPlfLst: number,
  fcPlfLfo: number,
  lcbPlfLfo: number,
): DocNumberingCatalog {
  const catalog: DocNumberingCatalog = {
    lists: new Map(),
    instanceListIds: [],
    counters: new Map(),
  };
  if (
    fcPlfLst === undefined ||
    !lcbPlfLst ||
    fcPlfLst < 0 ||
    fcPlfLst + lcbPlfLst > tableStream.length
  ) {
    return catalog;
  }

  const listCount = Math.max(0, readInt16(tableStream, fcPlfLst) ?? 0);
  const definitions: Array<{
    listId: number;
    levelCount: number;
  }> = [];
  for (let index = 0; index < listCount; index += 1) {
    const definitionOffset = fcPlfLst + 2 + index * 28;
    if (definitionOffset + 28 > fcPlfLst + lcbPlfLst) break;
    const listId = readInt32(tableStream, definitionOffset);
    if (listId === undefined) continue;
    const flags = tableStream[definitionOffset + 26] ?? 0;
    definitions.push({
      listId,
      levelCount: flags & 0x01 ? 1 : 9,
    });
  }

  // PlfLst 的 LVL 数据紧跟 LSTF 数组；lcbPlfLst 是整个结构长度，不能当作 LVL 起点。
  let levelOffset = fcPlfLst + 2 + definitions.length * 28;
  definitions.forEach((definition) => {
    const levels: DocListLevel[] = [];
    for (
      let levelIndex = 0;
      levelIndex < definition.levelCount;
      levelIndex += 1
    ) {
      const parsed = readListLevel(tableStream, levelOffset);
      if (!parsed) break;
      levels.push(parsed.level);
      levelOffset = parsed.nextOffset;
    }
    if (levels.length) catalog.lists.set(definition.listId, levels);
  });

  if (
    fcPlfLfo > 0 &&
    lcbPlfLfo >= 4 &&
    fcPlfLfo + lcbPlfLfo <= tableStream.length
  ) {
    const instanceCount = readUint32(tableStream, fcPlfLfo) ?? 0;
    for (let index = 0; index < instanceCount; index += 1) {
      const listId = readInt32(tableStream, fcPlfLfo + 4 + index * 16);
      if (listId === undefined) break;
      catalog.instanceListIds.push(listId);
    }
  }
  return catalog;
}

function toRoman(value: number) {
  const symbols: Array<[number, string]> = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let rest = Math.max(1, value);
  return symbols.reduce((result, [unit, symbol]) => {
    let nextResult = result;
    while (rest >= unit) {
      nextResult += symbol;
      rest -= unit;
    }
    return nextResult;
  }, '');
}

function formatCounter(value: number, format: number) {
  if (format === 1 || format === 2) {
    const roman = toRoman(value);
    return format === 2 ? roman.toLowerCase() : roman;
  }
  if (format === 3 || format === 4) {
    let result = '';
    let rest = Math.max(1, value);
    while (rest > 0) {
      rest -= 1;
      result = String.fromCharCode(65 + (rest % 26)) + result;
      rest = Math.floor(rest / 26);
    }
    return format === 4 ? result.toLowerCase() : result;
  }
  return String(value);
}

/** 推进列表实例状态并返回当前段落的编号前缀。 */
export function nextDocNumberPrefix(
  catalog: DocNumberingCatalog,
  listId: number,
  levelIndex: number,
) {
  if (listId <= 0) return undefined;
  const definitionId = catalog.instanceListIds[listId - 1];
  const levels = catalog.lists.get(definitionId);
  const safeLevel = Math.max(0, Math.min(8, levelIndex));
  const level = levels?.[safeLevel] ?? levels?.[0];
  if (!level) return undefined;
  if (level.format === 0x17) {
    return { text: level.text, suffix: level.suffix, style: level.style };
  }

  const counters = catalog.counters.get(listId) ?? [];
  counters[safeLevel] =
    counters[safeLevel] === undefined
      ? Math.max(1, level.start)
      : counters[safeLevel] + 1;
  counters.splice(safeLevel + 1);
  catalog.counters.set(listId, counters);

  const text = level.text.replace(/%([1-9])/g, (_match, rawIndex) => {
    const index = Number(rawIndex) - 1;
    const referencedLevel = levels?.[index] ?? level;
    return formatCounter(
      counters[index] ?? Math.max(1, referencedLevel.start),
      referencedLevel.format,
    );
  });
  return { text, suffix: level.suffix, style: level.style };
}
