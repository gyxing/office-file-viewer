/** 组件能够识别和预览的 Office 文件格式。 */
export type PreviewKind = 'pptx' | 'ppt' | 'xlsx' | 'xls' | 'docx' | 'doc';

/** 多种文件格式复用的预览界面族。 */
export type PreviewFamily = 'word' | 'spreadsheet' | 'presentation';

/** 格式在 Worker 中可使用的最高解析能力。 */
export type OfficeWorkerCapability = 'none' | 'materialized' | 'source';

/** MIME 类型与应补充文件扩展名的稳定映射。 */
export type OfficeMimeTypeDefinition = {
  /** 浏览器响应或 Blob 声明的 MIME 类型。 */
  mimeType: string;
  /** MIME 类型缺少文件名时使用的扩展名。 */
  extension: string;
};

/** 单种 Office 格式的稳定能力元数据。 */
export type OfficeFormatMetadata = {
  /** 解析结果使用的格式标识。 */
  kind: PreviewKind;
  /** 可映射到当前格式的文件扩展名。 */
  extensions: readonly string[];
  /** 可映射到当前格式的 MIME 类型。 */
  mimeTypes: readonly OfficeMimeTypeDefinition[];
  /** 当前格式复用的预览界面族。 */
  family: PreviewFamily;
  /** 当前格式在 Worker 中可使用的最高解析能力。 */
  workerCapability: OfficeWorkerCapability;
};

/** OOXML PowerPoint、宏启用演示文稿和模板扩展名。 */
const PPTX_EXTENSIONS = ['.pptx', '.pptm', '.potx'] as const;
/** 二进制 PowerPoint 文件扩展名。 */
const PPT_EXTENSIONS = ['.ppt'] as const;
/** OOXML Excel、宏启用工作簿和模板扩展名。 */
const XLSX_EXTENSIONS = ['.xlsx', '.xlsm', '.xltx'] as const;
/** 二进制 Excel 文件扩展名。 */
const XLS_EXTENSIONS = ['.xls'] as const;
/** OOXML Word、宏启用文档和模板扩展名。 */
const DOCX_EXTENSIONS = ['.docx', '.docm', '.dotx'] as const;
/** 二进制 Word 与 WPS 文字文件扩展名。 */
const DOC_EXTENSIONS = ['.doc', '.wps'] as const;

/** 当前只读解析会忽略宏代码的扩展名。 */
const MACRO_ENABLED_OFFICE_EXTENSIONS = new Set(['.docm', '.xlsm', '.pptm']);

/** 当前预览器支持解析的 Office 文件扩展名。 */
export const SUPPORTED_OFFICE_EXTENSIONS = [
  ...PPTX_EXTENSIONS,
  ...PPT_EXTENSIONS,
  ...XLSX_EXTENSIONS,
  ...XLS_EXTENSIONS,
  ...DOCX_EXTENSIONS,
  ...DOC_EXTENSIONS,
] as const;

/** 文件选择器与格式识别共用的扩展名白名单。 */
export const OFFICE_FILE_ACCEPT = SUPPORTED_OFFICE_EXTENSIONS.join(',');

/** 六种解析类型的扩展名、MIME、预览族和 Worker 能力唯一来源。 */
export const OFFICE_FORMAT_METADATA = {
  pptx: {
    kind: 'pptx',
    extensions: PPTX_EXTENSIONS,
    mimeTypes: [
      {
        mimeType:
          'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        extension: '.pptx',
      },
      {
        mimeType: 'application/vnd.ms-powerpoint.presentation.macroenabled.12',
        extension: '.pptm',
      },
      {
        mimeType:
          'application/vnd.openxmlformats-officedocument.presentationml.template',
        extension: '.potx',
      },
    ],
    family: 'presentation',
    workerCapability: 'source',
  },
  ppt: {
    kind: 'ppt',
    extensions: PPT_EXTENSIONS,
    mimeTypes: [
      { mimeType: 'application/vnd.ms-powerpoint', extension: '.ppt' },
    ],
    family: 'presentation',
    workerCapability: 'source',
  },
  xlsx: {
    kind: 'xlsx',
    extensions: XLSX_EXTENSIONS,
    mimeTypes: [
      {
        mimeType:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        extension: '.xlsx',
      },
      {
        mimeType: 'application/vnd.ms-excel.sheet.macroenabled.12',
        extension: '.xlsm',
      },
      {
        mimeType:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
        extension: '.xltx',
      },
    ],
    family: 'spreadsheet',
    workerCapability: 'source',
  },
  xls: {
    kind: 'xls',
    extensions: XLS_EXTENSIONS,
    mimeTypes: [{ mimeType: 'application/vnd.ms-excel', extension: '.xls' }],
    family: 'spreadsheet',
    workerCapability: 'source',
  },
  docx: {
    kind: 'docx',
    extensions: DOCX_EXTENSIONS,
    mimeTypes: [
      {
        mimeType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        extension: '.docx',
      },
      {
        mimeType: 'application/vnd.ms-word.document.macroenabled.12',
        extension: '.docm',
      },
      {
        mimeType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
        extension: '.dotx',
      },
    ],
    family: 'word',
    workerCapability: 'source',
  },
  doc: {
    kind: 'doc',
    extensions: DOC_EXTENSIONS,
    mimeTypes: [
      { mimeType: 'application/msword', extension: '.doc' },
      { mimeType: 'application/wps-office.wps', extension: '.wps' },
    ],
    family: 'word',
    workerCapability: 'source',
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

/** MIME 类型到扩展名的映射由格式元数据生成，避免输入层重复维护。 */
export const OFFICE_MIME_EXTENSION_MAP: Readonly<Record<string, string>> =
  Object.fromEntries(
    Object.values(OFFICE_FORMAT_METADATA).flatMap((metadata) =>
      metadata.mimeTypes.map(({ mimeType, extension }) => [
        mimeType.toLowerCase(),
        extension,
      ]),
    ),
  );

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

/** 返回文件名末尾可识别的 Office 扩展名。 */
export function getOfficeFileExtension(fileName: string): string | undefined {
  const lowerFileName = fileName.toLowerCase();
  return OFFICE_EXTENSION_DEFINITIONS.find(({ extension }) =>
    lowerFileName.endsWith(extension),
  )?.extension;
}

/** 判断当前文件是否属于只读解析但不会执行宏的 OOXML 格式。 */
export function isMacroEnabledOfficeFileName(fileName: string): boolean {
  const extension = getOfficeFileExtension(fileName);
  return (
    extension !== undefined && MACRO_ENABLED_OFFICE_EXTENSIONS.has(extension)
  );
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
