import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import { readXml } from '../../shared/ooxml/archive';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  parseXml,
} from '../../shared/ooxml/xml';

/** DOCX 段落引用的编号实例与层级。 */
export type DocxNumberingReference = {
  numId: string;
  level?: number;
};

type DocxNumberingLevel = {
  level: number;
  start: number;
  format: string;
  text: string;
  suffix: 'tab' | 'space' | 'nothing';
  fontFamily?: string;
};

/** DOCX 自动编号定义及解析过程中的计数状态。 */
export type DocxNumberingCatalog = {
  abstracts: Record<string, Record<number, DocxNumberingLevel>>;
  instances: Record<
    string,
    {
      abstractId: string;
      startOverrides: Record<number, number>;
      levelOverrides: Record<number, DocxNumberingLevel>;
    }
  >;
  counters: Record<string, number[]>;
};

const readVal = (node: Element | null | undefined) =>
  attr(node, 'w:val') ?? attr(node, 'val');

/** 从段落属性读取自动编号引用。 */
export function readDocxNumberingReference(
  pPr: Element | null | undefined,
): DocxNumberingReference | undefined {
  const numPr = childByLocalName(pPr, 'numPr');
  const numId = readVal(childByLocalName(numPr, 'numId'));
  if (!numId || numId === '0') return undefined;
  const rawLevel = readVal(childByLocalName(numPr, 'ilvl'));
  const level = rawLevel === undefined ? undefined : Number(rawLevel);
  return {
    numId,
    level: Number.isFinite(level) ? level : undefined,
  };
}

function readNumberingLevel(node: Element): DocxNumberingLevel | undefined {
  const rawLevel = attr(node, 'w:ilvl') ?? attr(node, 'ilvl');
  const level = Number(rawLevel);
  if (!Number.isInteger(level) || level < 0 || level > 8) return undefined;
  const start = Number(readVal(childByLocalName(node, 'start')) ?? 1);
  const suffix = readVal(childByLocalName(node, 'suff'));
  const fonts = childByLocalName(childByLocalName(node, 'rPr'), 'rFonts');
  return {
    level,
    start: Number.isFinite(start) ? start : 1,
    format: readVal(childByLocalName(node, 'numFmt')) ?? 'decimal',
    text: readVal(childByLocalName(node, 'lvlText')) ?? `%${level + 1}`,
    suffix:
      suffix === 'tab' || suffix === 'nothing' || suffix === 'space'
        ? suffix
        : 'tab',
    fontFamily:
      attr(fonts, 'w:ascii') ??
      attr(fonts, 'ascii') ??
      attr(fonts, 'w:hAnsi') ??
      attr(fonts, 'hAnsi') ??
      attr(fonts, 'w:eastAsia') ??
      attr(fonts, 'eastAsia'),
  };
}

/** 读取 numbering.xml，建立与具体文档内容无关的编号模型。 */
export function readDocxNumbering(
  entries: OfficeEntryMap,
): DocxNumberingCatalog {
  const catalog: DocxNumberingCatalog = {
    abstracts: {},
    instances: {},
    counters: {},
  };
  const xml = readXml(entries, 'word/numbering.xml');
  if (!xml) return catalog;
  const root = parseXml(xml).documentElement;

  childrenByLocalName(root, 'abstractNum').forEach((abstractNode) => {
    const abstractId =
      attr(abstractNode, 'w:abstractNumId') ??
      attr(abstractNode, 'abstractNumId');
    if (!abstractId) return;
    const levels: Record<number, DocxNumberingLevel> = {};
    childrenByLocalName(abstractNode, 'lvl').forEach((levelNode) => {
      const level = readNumberingLevel(levelNode);
      if (level) levels[level.level] = level;
    });
    catalog.abstracts[abstractId] = levels;
  });

  childrenByLocalName(root, 'num').forEach((numNode) => {
    const numId = attr(numNode, 'w:numId') ?? attr(numNode, 'numId');
    const abstractId = readVal(childByLocalName(numNode, 'abstractNumId'));
    if (!numId || !abstractId) return;
    const startOverrides: Record<number, number> = {};
    const levelOverrides: Record<number, DocxNumberingLevel> = {};
    childrenByLocalName(numNode, 'lvlOverride').forEach((overrideNode) => {
      const level = Number(
        attr(overrideNode, 'w:ilvl') ?? attr(overrideNode, 'ilvl'),
      );
      if (!Number.isInteger(level)) return;
      const start = Number(
        readVal(childByLocalName(overrideNode, 'startOverride')),
      );
      if (Number.isFinite(start)) startOverrides[level] = start;
      const overrideLevel = childByLocalName(overrideNode, 'lvl');
      if (overrideLevel) {
        const parsed = readNumberingLevel(overrideLevel);
        if (parsed) levelOverrides[level] = parsed;
      }
    });
    catalog.instances[numId] = {
      abstractId,
      startOverrides,
      levelOverrides,
    };
  });
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
    while (rest >= unit) {
      result += symbol;
      rest -= unit;
    }
    return result;
  }, '');
}

function formatCounter(value: number, format: string) {
  if (format === 'upperLetter' || format === 'lowerLetter') {
    let result = '';
    let rest = Math.max(1, value);
    while (rest > 0) {
      rest -= 1;
      result = String.fromCharCode(65 + (rest % 26)) + result;
      rest = Math.floor(rest / 26);
    }
    return format === 'lowerLetter' ? result.toLowerCase() : result;
  }
  if (format === 'upperRoman' || format === 'lowerRoman') {
    const result = toRoman(value);
    return format === 'lowerRoman' ? result.toLowerCase() : result;
  }
  return String(value);
}

/** 推进编号状态并返回当前段落应显示的编号前缀。 */
export function nextDocxNumberPrefix(
  reference: DocxNumberingReference,
  catalog: DocxNumberingCatalog,
):
  | {
      text: string;
      suffix: DocxNumberingLevel['suffix'];
      fontFamily?: string;
    }
  | undefined {
  const instance = catalog.instances[reference.numId];
  if (!instance) return undefined;
  const levels = catalog.abstracts[instance.abstractId];
  const levelIndex = Math.max(0, Math.min(8, reference.level ?? 0));
  const level = instance.levelOverrides[levelIndex] ?? levels?.[levelIndex];
  if (!level) return undefined;

  const counters = (catalog.counters[reference.numId] ??= []);
  const start = instance.startOverrides[levelIndex] ?? level.start ?? 1;
  counters[levelIndex] =
    counters[levelIndex] === undefined ? start : counters[levelIndex] + 1;
  counters.splice(levelIndex + 1);

  const text = level.text.replace(/%([1-9])/g, (_match, rawIndex) => {
    const index = Number(rawIndex) - 1;
    const referencedLevel =
      instance.levelOverrides[index] ?? levels?.[index] ?? level;
    const value =
      counters[index] ??
      instance.startOverrides[index] ??
      referencedLevel.start ??
      1;
    return formatCounter(value, referencedLevel.format);
  });
  return { text, suffix: level.suffix, fontFamily: level.fontFamily };
}
