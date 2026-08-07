// DOC 渲染样式转换集中保留在格式层，避免解析服务依赖 React 类型。
import type { CSSProperties } from 'react';
import type { DocTextStyle } from '../../services/doc/types';

// DOC 解析出的文本样式字段和 React CSS 字段基本一一对应，集中转换便于后续补充新属性。
/** 将 DOC 文本样式转换为 React CSS 属性。 */
export function docTextStyleToCss(style?: DocTextStyle): CSSProperties {
  if (!style) return {};

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
    fontFamily: style.fontFamily,
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
  },
): CSSProperties {
  const css = docTextStyleToCss(style);
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
