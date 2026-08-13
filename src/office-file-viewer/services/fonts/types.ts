/** 宿主可覆盖的字体别名表，键为源字体，值为有序回退字体。 */
export type OfficeFileViewerFontAliases = Readonly<
  Record<string, readonly string[]>
>;

/** 字体解析与缺失诊断配置。 */
export type OfficeFileViewerFontOptions = {
  /** 覆盖内置字体别名的宿主字体映射。 */
  aliases?: OfficeFileViewerFontAliases;
  /** 追加到每条字体链中的全局回退字体。 */
  fallbackFamilies?: readonly string[];
  /** 是否通过警告回调报告源字体缺失，默认开启。 */
  warnOnMissing?: boolean;
};

/** 字体链解析结果。 */
export type OfficeFontResolution = {
  /** 用于缺失诊断的首个源字体。 */
  requestedFamily?: string;
  /** 浏览器将依次尝试的规范化字体名称。 */
  candidates: readonly string[];
  /** 可直接写入 CSS font-family 的字体链。 */
  cssFamily?: string;
};

/** 渲染入口统一使用的字体链解析函数。 */
export type OfficeFontFamilyResolver = (
  requestedFamily?: string,
) => string | undefined;

/** PPTX 主题字体方案中与字体解析相关的字段。 */
export type OfficeThemeFontScheme = {
  /** 拉丁文字使用的主字体。 */
  majorFont?: string;
  /** 拉丁文字使用的次字体。 */
  minorFont?: string;
  /** 东亚文字使用的主字体。 */
  majorEastAsiaFont?: string;
  /** 东亚文字使用的次字体。 */
  minorEastAsiaFont?: string;
};
