import type { OfficeFileViewerFontAliases } from './types';

/** 常见 Office 字体在不同桌面系统中的本地回退链。 */
export const DEFAULT_OFFICE_FONT_ALIASES: OfficeFileViewerFontAliases = {
  Calibri: ['Arial', 'Helvetica'],
  Cambria: ['Georgia', 'Times New Roman'],
  宋体: ['SimSun', 'Songti SC', 'Noto Serif CJK SC'],
  SimSun: ['宋体', 'Songti SC', 'Noto Serif CJK SC'],
  仿宋_GB2312: ['FangSong', '仿宋', 'STFangsong', 'Noto Serif CJK SC'],
  仿宋: ['FangSong', 'STFangsong', 'Noto Serif CJK SC'],
  FangSong: ['仿宋', 'STFangsong', 'Noto Serif CJK SC'],
  黑体: ['SimHei', 'Heiti SC', 'Noto Sans CJK SC'],
  SimHei: ['黑体', 'Heiti SC', 'Noto Sans CJK SC'],
  微软雅黑: ['Microsoft YaHei', 'PingFang SC', 'Noto Sans CJK SC'],
  'Microsoft YaHei': ['微软雅黑', 'PingFang SC', 'Noto Sans CJK SC'],
  'MS Gothic': ['ＭＳ ゴシック', 'Meiryo', 'Yu Gothic', 'Noto Sans CJK JP'],
  'ＭＳ ゴシック': ['MS Gothic', 'Meiryo', 'Yu Gothic', 'Noto Sans CJK JP'],
  Meiryo: ['Yu Gothic', 'Noto Sans CJK JP'],
};

/** CSS 通用字体族不参与字体缺失诊断。 */
export const OFFICE_GENERIC_FONT_FAMILIES = new Set([
  'serif',
  'sans-serif',
  'monospace',
  'cursive',
  'fantasy',
  'system-ui',
  'ui-serif',
  'ui-sans-serif',
  'ui-monospace',
  'emoji',
  'math',
  'fangsong',
]);

/** 判断字体是否更适合衬线通用族，避免无字体环境落到错误字形类别。 */
export function inferOfficeGenericFontFamily(fontFamily?: string) {
  if (!fontFamily) return 'sans-serif';
  return /cambria|times|georgia|simsun|宋体|songti|mingliu|mincho|serif/i.test(
    fontFamily,
  )
    ? 'serif'
    : 'sans-serif';
}
