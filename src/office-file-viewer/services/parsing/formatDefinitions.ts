/** 组件能够识别和预览的 Office 文件格式。 */
export type PreviewKind = 'pptx' | 'ppt' | 'xlsx' | 'xls' | 'docx' | 'doc';

/** 多种文件格式复用的预览界面族。 */
export type PreviewFamily = 'word' | 'spreadsheet' | 'presentation';

/** 单种 Office 格式的稳定能力元数据。 */
export type OfficeFormatMetadata = {
  /** 解析结果使用的格式标识。 */
  kind: PreviewKind;
  /** 可映射到当前格式的文件扩展名。 */
  extensions: readonly string[];
  /** 当前格式复用的预览界面族。 */
  family: PreviewFamily;
  /** 当前格式是否支持在 Worker 中解析。 */
  supportsWorker: boolean;
};

/** OOXML PowerPoint 文件扩展名。 */
const PPTX_EXTENSIONS = ['.pptx'] as const;
/** 二进制 PowerPoint 文件扩展名。 */
const PPT_EXTENSIONS = ['.ppt'] as const;
/** OOXML Excel 文件扩展名。 */
const XLSX_EXTENSIONS = ['.xlsx'] as const;
/** 二进制 Excel 文件扩展名。 */
const XLS_EXTENSIONS = ['.xls'] as const;
/** OOXML Word 文件扩展名。 */
const DOCX_EXTENSIONS = ['.docx'] as const;
/** 二进制 Word 与 WPS 文字文件扩展名。 */
const DOC_EXTENSIONS = ['.doc', '.wps'] as const;

/** 当前预览器支持解析的 Office 文件扩展名。 */
export const SUPPORTED_OFFICE_EXTENSIONS = [
  ...PPTX_EXTENSIONS,
  ...PPT_EXTENSIONS,
  ...XLSX_EXTENSIONS,
  ...XLS_EXTENSIONS,
  ...DOCX_EXTENSIONS,
  ...DOC_EXTENSIONS,
] as const;

/** 六种格式的扩展名、预览族和 Worker 能力唯一来源。 */
export const OFFICE_FORMAT_METADATA = {
  pptx: {
    kind: 'pptx',
    extensions: PPTX_EXTENSIONS,
    family: 'presentation',
    supportsWorker: false,
  },
  ppt: {
    kind: 'ppt',
    extensions: PPT_EXTENSIONS,
    family: 'presentation',
    supportsWorker: true,
  },
  xlsx: {
    kind: 'xlsx',
    extensions: XLSX_EXTENSIONS,
    family: 'spreadsheet',
    supportsWorker: false,
  },
  xls: {
    kind: 'xls',
    extensions: XLS_EXTENSIONS,
    family: 'spreadsheet',
    supportsWorker: true,
  },
  docx: {
    kind: 'docx',
    extensions: DOCX_EXTENSIONS,
    family: 'word',
    supportsWorker: false,
  },
  doc: {
    kind: 'doc',
    extensions: DOC_EXTENSIONS,
    family: 'word',
    supportsWorker: true,
  },
} as const satisfies Record<PreviewKind, OfficeFormatMetadata>;

/** 扩展名按长度降序排列，避免未来的包含关系导致短后缀提前命中。 */
const OFFICE_EXTENSION_DEFINITIONS = Object.values(OFFICE_FORMAT_METADATA)
  .flatMap((metadata) =>
    metadata.extensions.map((extension) => ({
      extension,
      kind: metadata.kind,
    })),
  )
  .sort((left, right) => right.extension.length - left.extension.length);

/** 返回目标格式的稳定能力元数据。 */
export function getOfficeFormatMetadata(
  kind: PreviewKind,
): OfficeFormatMetadata {
  return OFFICE_FORMAT_METADATA[kind];
}

/** 返回目标格式复用的预览界面族。 */
export function getPreviewFamily(kind: PreviewKind): PreviewFamily {
  return getOfficeFormatMetadata(kind).family;
}

/** 严格按文件扩展名识别格式，无法识别时不执行历史回退。 */
export function tryDetectPreviewKind(
  fileName: string,
): PreviewKind | undefined {
  const lowerFileName = fileName.toLowerCase();
  return OFFICE_EXTENSION_DEFINITIONS.find(({ extension }) =>
    lowerFileName.endsWith(extension),
  )?.kind;
}
