import type { OfficeFontFamilyResolver } from '../../../services/fonts/types';

/** 合并拉丁与东亚字体槽，让浏览器按字符选择与 Office 一致的字形。 */
export function resolvePptxFontFamily(
  resolver: OfficeFontFamilyResolver,
  fontFamily?: string,
  eastAsiaFontFamily?: string,
) {
  const source = Array.from(
    new Set([fontFamily, eastAsiaFontFamily].filter(Boolean)),
  ).join(', ');
  return resolver(source || undefined);
}
