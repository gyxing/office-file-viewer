// DOC 渲染样式转换集中保留在格式层，避免解析服务依赖 React 类型。
import type { CSSProperties } from 'react';
import type { DocTextStyle } from '../../services/doc/types';
import type { OfficeFontFamilyResolver } from '../../services/fonts/types';

/** 需要显式东亚字体回退的文字与全角标点范围。 */
const DOC_CJK_TEXT_PATTERN = /[\u3000-\u303f\u3400-\u9fff\uff00-\uffef]/u;
/** DOC 字体栈中可识别的东亚字体名称。 */
const DOC_EAST_ASIA_FONT_PATTERN =
  /宋体|新宋体|黑体|楷体|仿宋|微软雅黑|simsun|nsimsun|simhei|kaiti|fangsong|yahei|songti|heiti|noto\s+(?:sans|serif)\s+cjk/i;

/** 中文文本没有显式东亚字体时补回旧版 Word 的默认宋体。 */
function withDocEastAsiaFontFallback(
  fontFamily: string | undefined,
  text: string | undefined,
) {
  if (
    !fontFamily ||
    !text ||
    !DOC_CJK_TEXT_PATTERN.test(text) ||
    DOC_EAST_ASIA_FONT_PATTERN.test(fontFamily)
  ) {
    return fontFamily;
  }
  const sourceFamilies = fontFamily
    .split(',')
    .map((font) => font.trim())
    .filter((font) => !/^(?:serif|sans-serif|monospace)$/i.test(font));
  return [...sourceFamilies, 'SimSun', '"宋体"'].join(', ');
}

// DOC 解析出的文本样式字段和 React CSS 字段基本一一对应，集中转换便于后续补充新属性。
/** 将 DOC 文本样式转换为 React CSS 属性。 */
export function docTextStyleToCss(
  style?: DocTextStyle,
  resolveFontFamily?: OfficeFontFamilyResolver,
  text?: string,
): CSSProperties {
  if (!style) return {};

  const fontFamily = withDocEastAsiaFontFallback(style.fontFamily, text);
  const css: CSSProperties = {
    color: style.color,
    background: style.backgroundColor,
    borderColor: style.borderColor,
    borderWidth: style.borderWidth,
    borderStyle: style.borderStyle,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight,
    fontStyle: style.fontStyle,
    textDecoration: style.textDecoration,
    textAlign: style.textAlign,
    // React 的数字 lineHeight 会按无单位倍数处理；DOC 解析出的绝对行距需显式补 px。
    lineHeight:
      style.lineHeight !== undefined && style.lineHeight > 4
        ? `${style.lineHeight}px`
        : style.lineHeight,
    fontFamily: resolveFontFamily?.(fontFamily) ?? fontFamily,
    marginLeft: style.indentLeft,
    marginRight: style.indentRight,
    textIndent: style.firstLineIndent,
    marginTop: style.spacingBefore,
    marginBottom: style.spacingAfter,
    paddingTop: style.paddingTop,
    paddingRight: style.paddingRight,
    paddingBottom: style.paddingBottom,
    paddingLeft: style.paddingLeft,
  };
  return Object.fromEntries(
    Object.entries(css).filter(([, value]) => value !== undefined),
  ) as CSSProperties;
}

/** 将 DOC 行内样式转换为 React CSS 属性。 */
export function inlineStyleToCss(
  style?: DocTextStyle,
  options?: {
    /** 是否保留块级模型自身的字体与段落样式。 */
    preserveBlockTypography?: boolean;
    /** 当前样式实际承载的文字，用于缺失东亚字体时选择回退。 */
    text?: string;
  },
  resolveFontFamily?: OfficeFontFamilyResolver,
): CSSProperties {
  const css = docTextStyleToCss(style, resolveFontFamily, options?.text);
  // 行内片段不能继承段落级缩进/间距，否则会把整段排版撑乱。
  delete css.textAlign;
  delete css.marginLeft;
  delete css.marginRight;
  delete css.textIndent;
  delete css.marginTop;
  delete css.marginBottom;
  delete css.paddingTop;
  delete css.paddingRight;
  delete css.paddingBottom;
  delete css.paddingLeft;

  if (options?.preserveBlockTypography) {
    delete css.fontSize;
    delete css.fontWeight;
    delete css.lineHeight;
  }

  return css;
}
