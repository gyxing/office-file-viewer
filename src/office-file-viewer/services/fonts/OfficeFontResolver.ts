import {
  DEFAULT_OFFICE_FONT_ALIASES,
  inferOfficeGenericFontFamily,
  OFFICE_GENERIC_FONT_FAMILIES,
} from './officeFontAliases';
import type {
  OfficeFileViewerFontOptions,
  OfficeFontResolution,
  OfficeThemeFontScheme,
} from './types';

/** PANOSE 第三个字节表示字重，按由细到粗映射到 CSS 的九档字重。 */
const PANOSE_WEIGHT_TO_CSS: Record<number, number> = {
  2: 100,
  3: 200,
  4: 300,
  5: 400,
  6: 500,
  7: 600,
  8: 700,
  9: 800,
  10: 900,
  11: 900,
};

/** 旧版 Office 未保存 PANOSE 时，用字体自身已知的基础字重补足缺失元数据。 */
const OFFICE_FONT_BASE_WEIGHT_HINTS: Readonly<Record<string, number>> = {
  'noto sans sc': 100,
};

/** 字体名称显式携带的字重后缀及其 CSS 字重。 */
const OFFICE_FONT_WEIGHT_SUFFIXES: ReadonlyArray<
  readonly [pattern: RegExp, weight: number]
> = [
  [/(?:thin|hairline)$/i, 100],
  [/(?:extra|ultra)[\s-]*light$/i, 200],
  [/(?:demi[\s-]*light|light)$/i, 300],
  [/(?:medium)$/i, 500],
  [/(?:semi|demi)[\s-]*bold$/i, 600],
  [/(?:extra|ultra)[\s-]*bold$/i, 800],
  [/(?:black|heavy)$/i, 900],
  [/(?:bold)$/i, 700],
];

/** 去除字体名外围引号，保留名称内部字符。 */
function normalizeFontFamily(fontFamily: string) {
  const trimmed = fontFamily.trim();
  if (
    trimmed.length >= 2 &&
    ((trimmed.startsWith('"') && trimmed.endsWith('"')) ||
      (trimmed.startsWith("'") && trimmed.endsWith("'")))
  ) {
    return trimmed.slice(1, -1).trim();
  }
  return trimmed;
}

/** 拆分 CSS 字体栈，同时避免误拆引号中的逗号。 */
function splitFontFamilies(fontFamily?: string) {
  if (!fontFamily) return [];
  const families: string[] = [];
  let current = '';
  let quote = '';
  for (const character of fontFamily) {
    if ((character === '"' || character === "'") && !quote) {
      quote = character;
      current += character;
    } else if (character === quote) {
      quote = '';
      current += character;
    } else if (character === ',' && !quote) {
      const normalized = normalizeFontFamily(current);
      if (normalized) families.push(normalized);
      current = '';
    } else {
      current += character;
    }
  }
  const normalized = normalizeFontFamily(current);
  if (normalized) families.push(normalized);
  return families;
}

/** 按字体名大小写不敏感地读取别名，宿主映射优先于内置映射。 */
function findFontAliases(
  fontFamily: string,
  options?: OfficeFileViewerFontOptions,
) {
  const normalized = fontFamily.toLocaleLowerCase();
  const customEntry = Object.entries(options?.aliases ?? {}).find(
    ([family]) => family.toLocaleLowerCase() === normalized,
  );
  if (customEntry) return customEntry[1];
  return Object.entries(DEFAULT_OFFICE_FONT_ALIASES).find(
    ([family]) => family.toLocaleLowerCase() === normalized,
  )?.[1];
}

/** 对包含空格或特殊字符的字体名补引号，通用族保持原样。 */
function formatCssFontFamily(fontFamily: string) {
  const normalized = fontFamily.toLocaleLowerCase();
  if (OFFICE_GENERIC_FONT_FAMILIES.has(normalized)) return normalized;
  if (/^[-_a-zA-Z0-9\u0080-\uFFFF]+$/u.test(fontFamily)) return fontFamily;
  return `"${fontFamily.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

/** 解析源字体与回退链。 */
export function resolveOfficeFont(
  requestedFamily?: string,
  options?: OfficeFileViewerFontOptions,
): OfficeFontResolution {
  const requestedCandidates = splitFontFamilies(requestedFamily);
  const requestedSource = requestedCandidates.find(
    (candidate) =>
      !OFFICE_GENERIC_FONT_FAMILIES.has(candidate.toLocaleLowerCase()),
  );
  const sourceCandidates = [
    ...requestedCandidates,
    ...(requestedSource ? findFontAliases(requestedSource, options) ?? [] : []),
    ...(options?.fallbackFamilies ?? []),
  ];
  const candidates: string[] = [];
  const generics: string[] = [];
  const seen = new Set<string>();
  sourceCandidates.forEach((candidate) => {
    splitFontFamilies(candidate).forEach((family) => {
      const identity = family.toLocaleLowerCase();
      if (seen.has(identity)) return;
      seen.add(identity);
      if (OFFICE_GENERIC_FONT_FAMILIES.has(identity)) {
        generics.push(identity);
      } else {
        candidates.push(family);
      }
    });
  });
  if (!candidates.length && !generics.length) {
    return { candidates: [] };
  }
  candidates.push(generics[0] ?? inferOfficeGenericFontFamily(requestedSource));
  return {
    requestedFamily: requestedSource,
    candidates,
    cssFamily: candidates.map(formatCssFontFamily).join(', '),
  };
}

/** 仅返回适合直接写入渲染样式的 CSS 字体链。 */
export function resolveOfficeFontFamily(
  requestedFamily?: string,
  options?: OfficeFileViewerFontOptions,
) {
  return resolveOfficeFont(requestedFamily, options).cssFamily;
}

/** 将 PPTX 主题字体标记还原为主题中声明的源字体。 */
export function resolveOfficeThemeFontFamily(
  typeface: string | undefined,
  fontScheme?: OfficeThemeFontScheme,
) {
  if (!typeface || !fontScheme) return typeface;
  const themeFonts: Record<string, string | undefined> = {
    '+mj-lt': fontScheme.majorFont,
    '+mn-lt': fontScheme.minorFont,
    '+mj-ea': fontScheme.majorEastAsiaFont ?? fontScheme.majorFont,
    '+mn-ea': fontScheme.minorEastAsiaFont ?? fontScheme.minorFont,
    '+mj-cs': fontScheme.majorFont,
    '+mn-cs': fontScheme.minorFont,
  };
  return themeFonts[typeface.toLocaleLowerCase()] ?? typeface;
}

/** 从 Office 字体声明的 PANOSE 值读取基础字重。 */
export function resolveOfficePanoseFontWeight(panose?: string) {
  if (!panose || !/^[0-9a-f]{6,}$/i.test(panose)) return undefined;
  const weightCode = Number.parseInt(panose.slice(4, 6), 16);
  return PANOSE_WEIGHT_TO_CSS[weightCode];
}

/** 在格式未保存字体字重元数据时，从字体名称和已知字体档案推断基础字重。 */
export function inferOfficeFontBaseWeight(fontFamily?: string) {
  const sourceFamily = splitFontFamilies(fontFamily)[0];
  if (!sourceFamily) return undefined;
  const normalized = sourceFamily.toLocaleLowerCase();
  const hintedWeight = OFFICE_FONT_BASE_WEIGHT_HINTS[normalized];
  if (hintedWeight !== undefined) return hintedWeight;
  return OFFICE_FONT_WEIGHT_SUFFIXES.find(([pattern]) =>
    pattern.test(sourceFamily),
  )?.[1];
}

/** 在保留源字体基础字重的前提下应用 Office 粗体修饰。 */
export function resolveOfficeCssFontWeight(
  baseWeight: number | undefined,
  bold: boolean | undefined,
) {
  if (baseWeight === undefined) return bold ? 700 : 400;
  if (!bold) return baseWeight;
  return Math.min(900, baseWeight + (baseWeight <= 400 ? 300 : 200));
}
