import type { OfficeSearchQuery } from './types';

/** 归一化文本及其每个 UTF-16 单元对应的原始字符区间。 */
export type NormalizedSearchText = Readonly<{
  /** 用于执行字面量匹配的归一化文本。 */
  text: string;
  /** 每个归一化文本单元在原文中的起始偏移。 */
  startOffsets: readonly number[];
  /** 每个归一化文本单元在原文中的结束偏移。 */
  endOffsets: readonly number[];
}>;

/** 搜索匹配在未经归一化的原始文本中的字符区间。 */
export type SearchTextMatch = Readonly<{
  startOffset: number;
  endOffset: number;
}>;

const INVISIBLE_CONTROL_PATTERN =
  /^[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f\u200b-\u200d\u2060\ufeff]$/u;
const HAN_PATTERN = /\p{Script=Han}/u;
const WORD_CHARACTER_PATTERN = /^[\p{L}\p{N}_]$/u;

function appendMappedText(
  target: string[],
  startOffsets: number[],
  endOffsets: number[],
  text: string,
  rawStart: number,
  rawEnd: number,
) {
  target.push(text);
  for (let index = 0; index < text.length; index += 1) {
    startOffsets.push(rawStart);
    endOffsets.push(rawEnd);
  }
}

/**
 * 归一化换行、不可见控制字符和大小写，同时保留回到原始文本的偏移映射。
 */
export function normalizeSearchText(
  source: string,
  matchCase: boolean,
): NormalizedSearchText {
  const output: string[] = [];
  const startOffsets: number[] = [];
  const endOffsets: number[] = [];

  for (let index = 0; index < source.length; ) {
    const rawStart = index;
    const codePoint = source.codePointAt(index);
    if (codePoint === undefined) break;
    let character = String.fromCodePoint(codePoint);
    index += character.length;

    if (character === '\r') {
      if (source[index] === '\n') index += 1;
      character = '\n';
    } else if (character === '\t') {
      character = ' ';
    } else if (INVISIBLE_CONTROL_PATTERN.test(character)) {
      continue;
    }

    const normalizedCharacter = matchCase
      ? character
      : character.toLocaleLowerCase();
    appendMappedText(
      output,
      startOffsets,
      endOffsets,
      normalizedCharacter,
      rawStart,
      index,
    );
  }

  return { text: output.join(''), startOffsets, endOffsets };
}

function codePointBefore(text: string, index: number) {
  if (index <= 0) return '';
  const previousUnit = text.charCodeAt(index - 1);
  const start =
    previousUnit >= 0xdc00 && previousUnit <= 0xdfff ? index - 2 : index - 1;
  const codePoint = text.codePointAt(Math.max(0, start));
  return codePoint === undefined ? '' : String.fromCodePoint(codePoint);
}

function codePointAt(text: string, index: number) {
  const value = text.codePointAt(index);
  return value === undefined ? '' : String.fromCodePoint(value);
}

function isWholeWordMatch(text: string, start: number, end: number) {
  const first = codePointAt(text, start);
  const last = codePointBefore(text, end);
  const before = codePointBefore(text, start);
  const after = codePointAt(text, end);
  return !(
    (WORD_CHARACTER_PATTERN.test(first) &&
      WORD_CHARACTER_PATTERN.test(before)) ||
    (WORD_CHARACTER_PATTERN.test(last) && WORD_CHARACTER_PATTERN.test(after))
  );
}

/** 使用字面量子串搜索并返回原始文本偏移，用户输入不会作为正则执行。 */
export function findSearchMatches(
  source: string,
  query: OfficeSearchQuery,
): SearchTextMatch[] {
  const haystack = normalizeSearchText(source, query.matchCase);
  const needle = normalizeSearchText(query.text, query.matchCase).text;
  if (!needle) return [];

  const enforceWordBoundary = query.wholeWord && !HAN_PATTERN.test(needle);
  const matches: SearchTextMatch[] = [];
  let fromIndex = 0;
  while (fromIndex <= haystack.text.length - needle.length) {
    const start = haystack.text.indexOf(needle, fromIndex);
    if (start < 0) break;
    const end = start + needle.length;
    if (!enforceWordBoundary || isWholeWordMatch(haystack.text, start, end)) {
      matches.push({
        startOffset: haystack.startOffsets[start],
        endOffset: haystack.endOffsets[end - 1],
      });
    }
    fromIndex = start + Math.max(1, needle.length);
  }
  return matches;
}

/** 截取匹配结果附近的上下文，避免结果列表复制整段长文本。 */
export function createSearchPreviewText(
  source: string,
  startOffset: number,
  endOffset: number,
  maximumLength = 96,
) {
  if (source.length <= maximumLength) return source;
  const matchLength = Math.max(0, endOffset - startOffset);
  const sideLength = Math.max(0, Math.floor((maximumLength - matchLength) / 2));
  const start = Math.max(0, startOffset - sideLength);
  const end = Math.min(source.length, endOffset + sideLength);
  return `${start > 0 ? '…' : ''}${source.slice(start, end)}${
    end < source.length ? '…' : ''
  }`;
}
