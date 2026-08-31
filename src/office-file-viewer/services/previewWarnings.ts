import { isMacroEnabledOfficeFileName } from './parsing/formatDefinitions';
import type { OfficeFileViewerPreviewState } from './parsing/internalTypes';
import type { PreviewKind } from './preview';

/** 区分解析、预览降级、链接与字体运行时产生的警告。 */
export type OfficeFileViewerWarningSource =
  | 'parser'
  | 'partial-preview'
  | 'hyperlink'
  | 'font';

/** 解析、预览降级或链接运行时产生的通用警告。 */
export type OfficeFileViewerGenericWarning = {
  /** 供宿主稳定识别警告类型的代码。 */
  code: string;
  /** 面向日志或界面展示的警告说明。 */
  message: string;
  /** 产生警告的文件格式。 */
  previewKind: PreviewKind;
  /** 警告来自解析器、部分预览降级或链接运行时。 */
  source: Exclude<OfficeFileViewerWarningSource, 'font'>;
};

/** 源文档字体不可用时产生的结构化回退警告。 */
export type OfficeFontFallbackWarning = {
  /** 供宿主稳定识别字体回退的代码。 */
  code: 'FONT_FALLBACK_APPLIED';
  /** 面向日志或界面展示的警告说明。 */
  message: string;
  /** 产生警告的文件格式。 */
  previewKind: PreviewKind;
  /** 固定标识字体运行时诊断。 */
  source: 'font';
  /** 源文档请求但当前浏览器不可用的字体。 */
  requestedFamily: string;
  /** 浏览器将依次尝试的完整字体链。 */
  candidates: readonly string[];
};

/** 宿主声明的字体资源加载失败时产生的结构化警告。 */
export type OfficeFontSourceLoadWarning = {
  /** 供宿主稳定识别字体资源加载失败。 */
  code: 'FONT_SOURCE_LOAD_FAILED';
  /** 面向日志或界面展示的警告说明。 */
  message: string;
  /** 产生警告的文件格式。 */
  previewKind: PreviewKind;
  /** 固定标识字体运行时诊断。 */
  source: 'font';
  /** 加载失败的字体族名称。 */
  requestedFamily: string;
  /** 失败后仍会尝试的字体链。 */
  candidates: readonly string[];
};

/** OfficeFileViewer 向宿主报告的非致命问题。 */
export type OfficeFileViewerWarning =
  | OfficeFileViewerGenericWarning
  | OfficeFontFallbackWarning
  | OfficeFontSourceLoadWarning;

/** 根据文件扩展名产生不依赖具体解析路径的只读安全警告。 */
export function collectOfficeFileWarnings(
  fileName: string,
  previewKind: PreviewKind,
): OfficeFileViewerWarning[] {
  if (!isMacroEnabledOfficeFileName(fileName)) return [];
  return [
    {
      code: 'MACRO_CONTENT_IGNORED',
      message: '当前以只读方式预览文档内容，宏代码不会加载或执行。',
      previewKind,
      source: 'parser',
    },
  ];
}
/** 从 DOC 字符串警告中恢复可选的稳定代码前缀。 */
function parseDocWarning(message: string) {
  const match = /^\[([A-Z0-9_]+)\]\s*(.*)$/s.exec(message);
  return match
    ? { code: match[1], message: match[2] }
    : { code: 'DOC_PARSE_WARNING', message };
}

/** 读取物化模型或按需数据源已经公开的解析警告。 */
export function collectOfficePreviewWarnings(
  preview: OfficeFileViewerPreviewState,
): OfficeFileViewerWarning[] {
  let warnings: ReadonlyArray<{ code: string; message: string }> = [];
  if (preview.mode === 'source') {
    if (preview.previewKind === 'doc') {
      warnings = preview.summary.warnings.map(parseDocWarning);
    } else if (preview.previewKind === 'docx') {
      warnings = preview.summary.review?.warnings ?? [];
    } else if (
      preview.previewKind === 'ppt' ||
      preview.previewKind === 'pptx'
    ) {
      warnings = preview.summary.warnings ?? [];
    }
  } else if (preview.model.kind === 'doc') {
    warnings = preview.model.document.warnings.map(parseDocWarning);
  } else if (preview.model.kind === 'docx') {
    warnings = preview.model.document.review?.warnings ?? [];
  } else if (preview.model.kind === 'ppt' || preview.model.kind === 'pptx') {
    warnings = preview.model.document.warnings ?? [];
  } else if (preview.model.kind === 'xls' || preview.model.kind === 'xlsx') {
    warnings = preview.model.workbook.warnings ?? [];
  }

  const seen = new Set<string>();
  return warnings.flatMap((warning) => {
    const identity = `${warning.code}\u0000${warning.message}`;
    if (seen.has(identity)) return [];
    seen.add(identity);
    return [
      {
        code: warning.code,
        message: warning.message,
        previewKind: preview.previewKind,
        source: 'parser' as const,
      },
    ];
  });
}
