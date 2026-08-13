import {
  attr,
  childByLocalName,
  childrenByLocalName,
  descendantByLocalName,
  parseXml,
} from './xml';

/** Office 主题。 */
export type OfficeTheme = {
  /** 按主题槽位索引的颜色方案。 */
  colorScheme: Record<string, string>;
  /** 将逻辑颜色槽映射到主题颜色槽的规则。 */
  colorMap?: Record<string, string>;
  /** 主题主字体和次字体方案。 */
  fontScheme?: Record<string, string>;
};

/** Office 主题颜色角色到颜色槽位的默认映射。 */
const DEFAULT_COLOR_MAP: Record<string, string> = {
  bg1: 'lt1',
  tx1: 'dk1',
  bg2: 'lt2',
  tx2: 'dk2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hlink',
  folHlink: 'folHlink',
};

/** 文档缺少主题定义时使用的默认 Office 主题。 */
export const DEFAULT_OFFICE_THEME: OfficeTheme = {
  colorMap: DEFAULT_COLOR_MAP,
  colorScheme: {
    dk1: '000000',
    lt1: 'FFFFFF',
    dk2: '44546A',
    lt2: 'E7E6E6',
    accent1: '5B9BD5',
    accent2: 'ED7D31',
    accent3: 'A5A5A5',
    accent4: 'FFC000',
    accent5: '4472C4',
    accent6: '70AD47',
    hlink: '0563C1',
    folHlink: '954F72',
  },
};

function toHexColor(value?: string) {
  if (!value) return undefined;
  if (value.startsWith('#')) return value;
  if (/^[0-9a-f]{6}$/i.test(value)) return `#${value}`;
  return undefined;
}

function readFontScheme(fontSchemeNode: Element | null | undefined) {
  const scheme: Record<string, string> = {};
  if (!fontSchemeNode) return scheme;

  ['majorFont', 'minorFont'].forEach((bucket) => {
    const node = childByLocalName(fontSchemeNode, bucket);
    if (!node) return;
    const latin = childByLocalName(node, 'latin');
    const ea = childByLocalName(node, 'ea');
    const cs = childByLocalName(node, 'cs');
    const scriptFonts = new Map(
      childrenByLocalName(node, 'font')
        .map((font) => [attr(font, 'script'), attr(font, 'typeface')] as const)
        .filter((entry): entry is readonly [string, string] =>
          Boolean(entry[0] && entry[1]),
        ),
    );
    const latinTypeface = attr(latin, 'typeface');
    const complexTypeface = attr(cs, 'typeface');
    const buildFontStack = (eastAsiaTypeface?: string) =>
      [eastAsiaTypeface, latinTypeface, complexTypeface]
        .filter(Boolean)
        .join(', ');
    // Office 主题同时保存多套东亚字体；缺少语言信息时优先使用简体中文槽位。
    const eastAsia =
      attr(ea, 'typeface') ||
      scriptFonts.get('Hans') ||
      scriptFonts.get('Hant') ||
      scriptFonts.get('Jpan') ||
      scriptFonts.get('Hang');
    const value = buildFontStack(eastAsia);
    if (value) scheme[bucket] = value;
    ['Hans', 'Hant', 'Jpan', 'Hang'].forEach((script) => {
      const scriptValue = buildFontStack(scriptFonts.get(script));
      if (scriptValue) scheme[`${bucket}:${script}`] = scriptValue;
    });
  });

  return scheme;
}

/** 读取 OOXML 主题或主题覆盖中的颜色和字体配置。 */
export function readOfficeTheme(
  xml?: string,
  baseTheme: OfficeTheme = DEFAULT_OFFICE_THEME,
): OfficeTheme {
  if (!xml) return baseTheme;

  const doc = parseXml(xml);
  const colorScheme: Record<string, string> = {
    ...baseTheme.colorScheme,
  };
  const themeElements = childByLocalName(doc.documentElement, 'themeElements');
  // 图表 themeOverride 直接包含 clrScheme/fontScheme，不再包一层 themeElements。
  const clrScheme =
    childByLocalName(themeElements, 'clrScheme') ??
    descendantByLocalName(doc.documentElement, 'clrScheme');
  const fontScheme = {
    ...(baseTheme.fontScheme ?? {}),
    ...readFontScheme(
      childByLocalName(themeElements, 'fontScheme') ??
        descendantByLocalName(doc.documentElement, 'fontScheme'),
    ),
  };

  Object.keys(DEFAULT_OFFICE_THEME.colorScheme).forEach((name) => {
    const node = childByLocalName(clrScheme, name);
    const child = node?.firstElementChild;
    const childName = child?.localName.split(':').pop()?.toLowerCase();
    const value =
      childName === 'sysclr'
        ? attr(child, 'lastClr') ?? attr(child, 'val')
        : attr(child, 'val') ?? attr(child, 'lastClr');
    if (value) colorScheme[name] = value;
  });

  return {
    colorScheme,
    colorMap: baseTheme.colorMap ?? DEFAULT_COLOR_MAP,
    fontScheme,
  };
}

/** 解析并确定 `resolveOfficeThemeColor` 对应的引用或配置。 */
export function resolveOfficeThemeColor(
  value: string | undefined,
  theme: OfficeTheme = DEFAULT_OFFICE_THEME,
) {
  if (!value) return undefined;

  const direct = toHexColor(value);
  if (direct) return direct;

  const mapped = theme.colorMap?.[value] ?? value;
  return toHexColor(theme.colorScheme[mapped] ?? theme.colorScheme[value]);
}
