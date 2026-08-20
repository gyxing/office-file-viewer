// docxRenderUtils 提供 DOCX 渲染阶段的样式转换和空段落高度计算。
import type { CSSProperties } from 'react';
import type {
  DocxParagraphBlock,
  DocxTextStyle,
} from '../../services/docx/types';
import type { OfficeFontFamilyResolver } from '../../services/fonts/types';

// DOCX 的样式已经在解析阶段完成继承合并，这里只负责把最终样式映射到 React CSS。

/** 可从字体名称识别的常见东亚字体，供脚本提示调整字体链优先级。 */
const DOCX_EAST_ASIA_FONT_PATTERN =
  /[\u3040-\u30ff\u3400-\u9fff\uac00-\ud7af]|simsun|simhei|yahei|dengxian|fangsong|kaiti|songti|heiti|mingliu|meiryo|malgun|noto\s+(?:sans|serif)\s+cjk|ms\s+gothic/i;

/** 判断当前文字是否按东亚脚本使用宋体字面。 */
function usesEastAsiaSimSun(style?: DocxTextStyle) {
  return Boolean(
    /(?:^|,)\s*["']?(?:宋体|simsun)["']?\s*(?:,|$)/i.test(
      style?.fontFamily ?? '',
    ),
  );
}

/** Office 形状文本中，微软雅黑粗体相对浏览器字体场景的基线补偿比例。 */
const DOCX_SHAPE_YAHEI_BASELINE_OFFSET_EM = 1 / 3;
/** 可按字体度量补偿形状文本基线的微软雅黑字体。 */
const DOCX_SHAPE_YAHEI_FONT_PATTERN = /微软雅黑|Microsoft\s+YaHei/i;
/** Word 中 11 磅字号换算到浏览器后的标准像素值。 */
const DOCX_SIMSUN_11PT_SIZE_PX = 44 / 3;

/** 抵消浏览器对 11 磅宋体字宽整像素吸附造成的累计换行误差。 */
const DOCX_SIMSUN_11PT_LETTER_SPACING_PX = -0.31;
/** Windows 小字号宋体原生合成粗体增加的单字前进宽度。 */
const DOCX_SIMSUN_HINTED_BOLD_ADVANCE_PX = 1;
/** 宋体关闭整数像素 hinting 后，原生合成粗体增加的字号比例。 */
const DOCX_SIMSUN_NATIVE_BOLD_ADVANCE_RATIO = 1 / 51.2;
/** Windows 宋体仍使用整数像素 hinting 的最大字号。 */
const DOCX_SIMSUN_HINTED_BOLD_MAX_SIZE_PX = 16;

/** 保留源字符间距，并抵消浏览器宋体 hinting 与原生合成粗体造成的宽度漂移。 */
function resolveSimSunLetterSpacing(style?: DocxTextStyle) {
  if (!usesEastAsiaSimSun(style)) return style?.letterSpacing;
  const fontSize = style?.fontSize;
  let letterSpacing = style?.letterSpacing ?? 0;
  if (
    style?.letterSpacing === undefined &&
    fontSize !== undefined &&
    Math.abs(fontSize - DOCX_SIMSUN_11PT_SIZE_PX) <= 0.01
  ) {
    letterSpacing += DOCX_SIMSUN_11PT_LETTER_SPACING_PX;
  }
  if (style?.bold && fontSize !== undefined) {
    letterSpacing -=
      fontSize <= DOCX_SIMSUN_HINTED_BOLD_MAX_SIZE_PX
        ? DOCX_SIMSUN_HINTED_BOLD_ADVANCE_PX
        : fontSize * DOCX_SIMSUN_NATIVE_BOLD_ADVANCE_RATIO;
  }
  return Math.abs(letterSpacing) > 0.0001 ? letterSpacing : undefined;
}

/** 按 OOXML 脚本提示调整字体链，避免系统字体链接改变东亚文字字宽。 */
export function resolveDocxTextFontFamily(
  style?: DocxTextStyle,
  resolveFontFamily?: OfficeFontFamilyResolver,
) {
  const sourceFamily = style?.fontFamily;
  if (!sourceFamily) return undefined;
  const families = sourceFamily
    .split(',')
    .map((family) => family.trim())
    .filter(Boolean);
  if (style.fontHint === 'eastAsia') {
    const eastAsiaIndex = families.findIndex((family) =>
      DOCX_EAST_ASIA_FONT_PATTERN.test(family),
    );
    if (eastAsiaIndex > 0) {
      families.unshift(...families.splice(eastAsiaIndex, 1));
    }
  }
  const orderedFamily = families.join(', ');
  return resolveFontFamily?.(orderedFamily) ?? orderedFamily;
}

/** 为东亚字体运行恢复 OOXML 单独声明的西文字体链。 */
export function resolveDocxLatinFontFamily(
  style?: DocxTextStyle,
  resolveFontFamily?: OfficeFontFamilyResolver,
) {
  if (style?.fontHint !== 'eastAsia' || !style.fontFamily) return undefined;
  const families = style.fontFamily
    .split(',')
    .map((family) => family.trim())
    .filter(Boolean);
  const eastAsiaIndex = families.findIndex((family) =>
    DOCX_EAST_ASIA_FONT_PATTERN.test(family),
  );
  if (eastAsiaIndex < 0 || families.length < 2) return undefined;
  const latinFamily = families
    .filter((_, index) => index !== eastAsiaIndex)
    .join(', ');
  return resolveFontFamily?.(latinFamily) ?? latinFamily;
}

/** 还原 Office 形状中微软雅黑粗体使用的字体场景基线。 */
export function resolveDocxShapeBaselineOffset(style?: DocxTextStyle) {
  if (
    !style?.bold ||
    style.fontSize === undefined ||
    !DOCX_SHAPE_YAHEI_FONT_PATTERN.test(style.fontFamily ?? '')
  ) {
    return undefined;
  }
  return style.fontSize * DOCX_SHAPE_YAHEI_BASELINE_OFFSET_EM;
}

/** 将 DOCX 文本样式转换为 React CSS 属性。 */
export function buildDocxTextStyle(
  style?: DocxTextStyle,
  options?: {
    /** 是否将背景样式写入渲染结果。 */
    includeBackground?: boolean;
  },
  resolveFontFamily?: OfficeFontFamilyResolver,
): CSSProperties {
  const css: CSSProperties = {
    // 缺少独立粗体字面时交给浏览器原生合成，避免描边造成文字重影。
    fontWeight:
      style?.bold === true ? 700 : style?.bold === false ? 400 : undefined,
    fontStyle:
      style?.italic === true
        ? 'italic'
        : style?.italic === false
        ? 'normal'
        : undefined,
    textDecoration:
      [style?.underline ? 'underline' : '', style?.strike ? 'line-through' : '']
        .filter(Boolean)
        .join(' ') || undefined,
    color: style?.color,
    fontSize: style?.fontSize,
    fontFamily: resolveDocxTextFontFamily(style, resolveFontFamily),
    letterSpacing: resolveSimSunLetterSpacing(style),

    textTransform: style?.allCaps ? 'uppercase' : undefined,
    fontVariant: style?.smallCaps ? 'small-caps' : undefined,
    background: options?.includeBackground ? style?.backgroundColor : undefined,
  };
  return Object.fromEntries(
    Object.entries(css).filter(([, value]) => value !== undefined),
  ) as CSSProperties;
}

/** 获取 `getDocxEmptyParagraphHeight` 返回的数据。 */
export function getDocxEmptyParagraphHeight(block: DocxParagraphBlock) {
  const fontSize = block.style?.fontSize ?? 14;
  // 空段落没有文字撑高，需要按行高补出可见高度，否则换行会丢失。
  if (block.lineHeight === undefined) return fontSize * 1.2;
  return block.lineHeight > 4 ? block.lineHeight : fontSize * block.lineHeight;
}

/** 将 DOCX 行高转换为 CSS 可识别的倍率或像素值。 */
export function getDocxCssLineHeight(block: DocxParagraphBlock) {
  if (block.lineHeight === undefined) return undefined;
  // OOXML 的 exact/atLeast 行距已经换算为像素，不能作为无单位倍率传给 React。
  return block.lineHeight > 4 ? `${block.lineHeight}px` : block.lineHeight;
}

/**
 * 读取分页测量使用的流式块高度，跳过修订投影后不再参与排版的隐藏块。
 */
export function measureVisibleDocxBlockHeight(
  elements: readonly HTMLElement[],
  index: number,
): number {
  const element = elements[index];
  const style = window.getComputedStyle(element);
  if (style.display === 'none' || element.offsetHeight === 0) return 0;

  // 隐藏修订块的 offsetTop 会回落到 0，必须寻找下一个可见块才能保持流式高度连续。
  for (let nextIndex = index + 1; nextIndex < elements.length; nextIndex += 1) {
    const candidate = elements[nextIndex];
    const candidateStyle = window.getComputedStyle(candidate);
    if (candidateStyle.display === 'none' || candidate.offsetHeight === 0) {
      continue;
    }
    return Math.max(0, candidate.offsetTop - element.offsetTop);
  }
  return element.offsetHeight + Number.parseFloat(style.marginBottom || '0');
}
