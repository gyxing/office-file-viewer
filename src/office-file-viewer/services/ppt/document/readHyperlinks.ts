import type { OfficeHyperlink } from '../../../shared/hyperlink';
import { createOfficeHyperlinkFromTarget } from '../../../shared/hyperlink';
import { PPT_RECORD } from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import { readPptUnicodeString } from '../binary/readStrings';
import type {
  PptHyperlinkEntry,
  PptParseContext,
  PptRecord,
  PptSlideDescriptor,
} from '../types';

/** InteractiveInfoAtom 中实际会触发导航的动作编号。 */
const PPT_ACTION = {
  jump: 0x03,
  hyperlink: 0x04,
} as const;

/** InteractiveInfoAtom 中的相对幻灯片跳转编号。 */
const PPT_JUMP_ACTIONS: Partial<
  Record<number, 'next' | 'previous' | 'first' | 'last'>
> = {
  0x01: 'next',
  0x02: 'previous',
  0x03: 'first',
  0x04: 'last',
};

/** LinkToEnum 中可直接映射到当前演示文稿的动作编号。 */
const PPT_LINK_ACTIONS: Partial<
  Record<number, 'next' | 'previous' | 'first' | 'last'>
> = {
  0x00: 'next',
  0x01: 'previous',
  0x02: 'first',
  0x03: 'last',
};

/** PPT 文本中由一条 InteractiveInfo 声明的链接范围。 */
export type PptTextHyperlinkRange = {
  /** 链接所属文本组的零基索引。 */
  groupIndex: number;
  /** 链接范围的起始字符位置。 */
  begin: number;
  /** 链接范围的截止字符位置。 */
  end: number;
  /** 统一后的链接目标。 */
  hyperlink: OfficeHyperlink;
};

function readRecordChildren(record: PptRecord) {
  if (record.version !== 0x0f) return [];
  return Array.from(new PptRecordReader(record.data).records());
}

function readHyperlinkEntry(record: PptRecord): PptHyperlinkEntry | undefined {
  const children = readRecordChildren(record);
  const idAtom = children.find(
    (child) =>
      child.type === PPT_RECORD.EXTERNAL_HYPERLINK_ATOM && child.length >= 4,
  );
  if (!idAtom) return undefined;
  const id = new DataView(
    idAtom.data.buffer,
    idAtom.data.byteOffset,
    idAtom.data.byteLength,
  ).getUint32(0, true);
  if (!id) return undefined;

  const strings = children
    .filter((child) => child.type === PPT_RECORD.C_STRING)
    .map((child) => ({
      instance: child.instance,
      value: readPptUnicodeString(child.data),
    }));
  const friendlyName =
    strings.find((item) => item.instance === 0)?.value ?? strings[0]?.value;
  const target =
    strings.find((item) => item.instance === 1)?.value ?? strings[1]?.value;
  const location =
    strings.find((item) => item.instance === 2)?.value ?? strings[2]?.value;
  return { id, friendlyName, target, location };
}

function collectHyperlinkEntries(
  record: PptRecord,
  entries: PptHyperlinkEntry[],
) {
  for (const child of readRecordChildren(record)) {
    if (child.type === PPT_RECORD.EXTERNAL_HYPERLINK) {
      const entry = readHyperlinkEntry(child);
      if (entry) entries.push(entry);
      continue;
    }
    if (child.version === 0x0f) collectHyperlinkEntries(child, entries);
  }
}

function reportPptHyperlinkWarning(context: PptParseContext, message: string) {
  if (
    context.warnings.some(
      (warning) =>
        warning.code === 'UNSUPPORTED_HYPERLINK' && warning.message === message,
    )
  ) {
    return;
  }
  context.warnings.push({ code: 'UNSUPPORTED_HYPERLINK', message });
}

/** 读取根 DocumentContainer 中的链接表，并缓存幻灯片标识映射。 */
export function readPptHyperlinks(
  documentRecord: PptRecord,
  descriptors: readonly PptSlideDescriptor[],
  context: PptParseContext,
) {
  context.hyperlinks.clear();
  context.slideIndexById.clear();
  descriptors.forEach((descriptor) => {
    context.slideIndexById.set(descriptor.slideId, descriptor.index);
  });

  const entries: PptHyperlinkEntry[] = [];
  try {
    collectHyperlinkEntries(documentRecord, entries);
  } catch {
    // 链接表损坏不应阻止正文、图片和版式继续恢复。
    reportPptHyperlinkWarning(context, 'PPT 文档中的超链接表无法完整读取');
  }
  entries.forEach((entry) => context.hyperlinks.set(entry.id, entry));
}

function createPresentationAction(
  action: 'next' | 'previous' | 'first' | 'last',
): OfficeHyperlink {
  return {
    kind: 'internal',
    target: { family: 'presentation', action },
  };
}

function parseSlideIndex(
  entry: PptHyperlinkEntry | undefined,
  context: PptParseContext,
) {
  const candidates = [entry?.location, entry?.target].filter(
    (value): value is string => Boolean(value),
  );
  for (const value of candidates) {
    const match = /^\s*(-?\d+)\s*,\s*(-?\d+)/.exec(value);
    if (match) {
      const slideId = Number(match[1]);
      const slideNumber = Number(match[2]);
      const mappedIndex = context.slideIndexById.get(slideId);
      // PPT 描述符保留源文件的一基页序号，统一导航模型使用零基索引。
      if (mappedIndex !== undefined) return mappedIndex - 1;
      if (slideNumber > 0) return slideNumber - 1;
    }
    const namedSlide = /(?:slide|幻灯片)\s*(\d+)/i.exec(value);
    if (namedSlide) return Number(namedSlide[1]) - 1;
  }
  return undefined;
}

function combineTarget(entry: PptHyperlinkEntry | undefined) {
  const target = entry?.target?.trim();
  const location = entry?.location?.trim();
  if (!target) return location;
  if (!location || target.includes('#')) return target;
  return `${target}#${location}`;
}

function parseInteractiveInfoAtom(
  atom: PptRecord,
  context: PptParseContext,
): OfficeHyperlink | undefined {
  if (atom.length < 16) return undefined;
  const view = new DataView(
    atom.data.buffer,
    atom.data.byteOffset,
    atom.data.byteLength,
  );
  const hyperlinkId = view.getUint32(4, true);
  const action = view.getUint8(8);
  const jump = view.getUint8(10);
  const hyperlinkType = view.getUint8(12);

  if (action === PPT_ACTION.jump) {
    const targetAction = PPT_JUMP_ACTIONS[jump];
    return targetAction ? createPresentationAction(targetAction) : undefined;
  }
  if (action !== PPT_ACTION.hyperlink) return undefined;

  const relativeAction = PPT_LINK_ACTIONS[hyperlinkType];
  if (relativeAction) return createPresentationAction(relativeAction);

  const entry = context.hyperlinks.get(hyperlinkId);
  if (hyperlinkType === 0x07) {
    const slideIndex = parseSlideIndex(entry, context);
    if (slideIndex === undefined) {
      reportPptHyperlinkWarning(context, 'PPT 内部幻灯片链接缺少有效目标');
      return undefined;
    }
    return {
      kind: 'internal',
      target: { family: 'presentation', slideIndex },
      screenTip: entry?.friendlyName,
    };
  }
  if (
    hyperlinkType !== 0x08 &&
    hyperlinkType !== 0x09 &&
    hyperlinkType !== 0x0a
  )
    return undefined;
  const target = combineTarget(entry);
  if (!target) {
    reportPptHyperlinkWarning(context, 'PPT 外部超链接缺少有效目标');
    return undefined;
  }
  return createOfficeHyperlinkFromTarget(target, entry?.friendlyName);
}

/** 从鼠标单击 InteractiveInfoContainer 恢复对象或文字链接。 */
export function readPptInteractiveHyperlink(
  record: PptRecord,
  context: PptParseContext,
) {
  if (
    record.type !== PPT_RECORD.INTERACTIVE_INFO ||
    record.instance !== 0 ||
    record.version !== 0x0f
  ) {
    return undefined;
  }
  try {
    const atom = readRecordChildren(record).find(
      (child) => child.type === PPT_RECORD.INTERACTIVE_INFO_ATOM,
    );
    return atom ? parseInteractiveInfoAtom(atom, context) : undefined;
  } catch {
    // 单个交互记录损坏时仅忽略该链接，形状和文字本身继续渲染。
    reportPptHyperlinkWarning(context, 'PPT 交互超链接记录无法读取');
    return undefined;
  }
}

/** 读取一个文本框内紧邻 InteractiveInfo 的字符链接范围。 */
export function readPptTextHyperlinkRanges(
  records: readonly PptRecord[],
  context: PptParseContext,
) {
  const ranges: PptTextHyperlinkRange[] = [];
  let groupIndex = -1;
  for (let index = 0; index < records.length; index += 1) {
    const record = records[index];
    if (
      record.type === PPT_RECORD.TEXT_CHARS_ATOM ||
      record.type === PPT_RECORD.TEXT_BYTES_ATOM
    ) {
      groupIndex += 1;
      continue;
    }
    const hyperlink = readPptInteractiveHyperlink(record, context);
    const range = records[index + 1];
    if (
      !hyperlink ||
      groupIndex < 0 ||
      range?.type !== PPT_RECORD.TEXT_INTERACTIVE_INFO_ATOM ||
      range.instance !== 0 ||
      range.length < 8
    ) {
      continue;
    }
    const view = new DataView(
      range.data.buffer,
      range.data.byteOffset,
      range.data.byteLength,
    );
    const begin = view.getUint32(0, true);
    const end = view.getUint32(4, true);
    if (end > begin) ranges.push({ groupIndex, begin, end, hyperlink });
    index += 1;
  }
  return ranges;
}
