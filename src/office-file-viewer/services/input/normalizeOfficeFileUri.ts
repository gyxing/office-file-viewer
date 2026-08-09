import type { OfficeFileViewerMessages } from '../../locale';
import { isSupportedOfficeFileName } from '../parsing/detectPreviewKind';

/** 异步文件来源可以消费取消信号；现有无参数函数仍保持类型兼容。 */
export type OfficeFileViewerUriLoader = (
  signal?: AbortSignal,
) => Promise<File | Blob | string | Response>;

/** 定义预览文件来源，可直接传入 File、URL，或返回文件数据的异步加载函数。 */
export type OfficeFileViewerUri = File | string | OfficeFileViewerUriLoader;

/** URI 输入归一化后交给控制器的文件及可靠远程来源。 */
export type NormalizedOfficeFileUri = {
  /** 可直接交给格式识别和解析会话的文件。 */
  file: File;
  /** 可用于解析文档相对链接的 HTTP(S) 来源地址。 */
  sourceUrl?: string;
};

/** Office MIME 类型到文件扩展名的映射。 */
const OFFICE_MIME_EXTENSION_MAP: Record<string, string> = {
  'application/vnd.openxmlformats-officedocument.presentationml.presentation':
    '.pptx',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
  'application/vnd.ms-excel': '.xls',
  'application/vnd.ms-powerpoint': '.ppt',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
    '.docx',
  'application/msword': '.doc',
  'application/wps-office.wps': '.wps',
};

/** 从 URL 路径末段提取并解码文件名；无法解析时返回 undefined。 */
function getFileNameFromUrl(url: string): string | undefined {
  try {
    const parsedUrl = new URL(url, window.location.href);
    const lastSegment = parsedUrl.pathname.split('/').filter(Boolean).pop();
    return lastSegment ? decodeURIComponent(lastSegment) : undefined;
  } catch {
    const path = url.split(/[?#]/)[0];
    const lastSegment = path.split('/').filter(Boolean).pop();
    return lastSegment ? decodeURIComponent(lastSegment) : undefined;
  }
}

/** 从 Content-Disposition 响应头提取 UTF-8 或普通文件名。 */
function getFileNameFromContentDisposition(
  contentDisposition: string | null,
): string | undefined {
  if (!contentDisposition) return undefined;

  const encodedMatch = contentDisposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (encodedMatch?.[1]) return decodeURIComponent(encodedMatch[1]);

  const plainMatch = contentDisposition.match(/filename="?([^";]+)"?/i);
  return plainMatch?.[1] ? decodeURIComponent(plainMatch[1]) : undefined;
}

/** 根据响应 MIME 类型推断受支持的 Office 文件扩展名。 */
function getExtensionFromMimeType(mimeType: string): string | undefined {
  return OFFICE_MIME_EXTENSION_MAP[
    mimeType.split(';')[0]?.trim().toLowerCase()
  ];
}

/** 判断文件名末尾是否包含扩展名。 */
function hasFileExtension(fileName: string): boolean {
  return /\.[^./\\]+$/.test(fileName);
}

/** 标记由预览器输入校验产生、可安全直接展示给用户的错误。 */
export class OfficeFileViewerInputError extends Error {}

/** 校验文件扩展名是否属于当前组件支持的 Office 格式。 */
export function ensureSupportedOfficeFile(
  file: File,
  messages: OfficeFileViewerMessages,
): void {
  if (!isSupportedOfficeFileName(file.name)) {
    throw new OfficeFileViewerInputError(messages.file.unsupported);
  }
}

/** 根据 Blob、显式文件名或 MIME 类型创建可识别格式的 File。 */
function createFileFromBlob(
  blob: Blob,
  messages: OfficeFileViewerMessages,
  fileName?: string,
): File {
  if (
    fileName &&
    hasFileExtension(fileName) &&
    !isSupportedOfficeFileName(fileName)
  ) {
    throw new OfficeFileViewerInputError(messages.file.unrecognized);
  }

  const extension = getExtensionFromMimeType(blob.type);
  const inferredFileName =
    fileName && isSupportedOfficeFileName(fileName)
      ? fileName
      : extension
      ? `office-file${extension}`
      : undefined;

  if (!inferredFileName) {
    throw new OfficeFileViewerInputError(messages.file.unrecognized);
  }

  return new File([blob], inferredFileName, { type: blob.type });
}

/** 校验下载响应，并将响应体及响应头文件名转换为 File。 */
async function createFileFromResponse(
  response: Response,
  messages: OfficeFileViewerMessages,
  fallbackFileName?: string,
): Promise<File> {
  if (!response.ok) {
    throw new OfficeFileViewerInputError(
      messages.file.downloadFailed(response.status, response.statusText),
    );
  }

  const blob = await response.blob();
  const fileName =
    getFileNameFromContentDisposition(
      response.headers.get('Content-Disposition'),
    ) || fallbackFileName;
  return createFileFromBlob(blob, messages, fileName);
}

/** 下载远程 Office 文件，并沿用 URL 或响应头中可识别的文件名。 */
async function downloadOfficeFile(
  url: string,
  messages: OfficeFileViewerMessages,
  signal?: AbortSignal,
): Promise<File> {
  const urlFileName = getFileNameFromUrl(url);
  if (
    urlFileName &&
    hasFileExtension(urlFileName) &&
    !isSupportedOfficeFileName(urlFileName)
  ) {
    throw new OfficeFileViewerInputError(messages.file.unsupported);
  }

  const response = await fetch(url, { signal });
  return createFileFromResponse(response, messages, urlFileName);
}

/** 仅保留可作为相对链接基准的 HTTP(S) 地址。 */
function getRemoteSourceUrl(value: string | undefined) {
  if (!value) return undefined;
  try {
    const baseUrl =
      typeof window === 'undefined' ? undefined : window.location.href;
    const url = baseUrl ? new URL(value, baseUrl) : new URL(value);
    return url.protocol === 'http:' || url.protocol === 'https:'
      ? url.href
      : undefined;
  } catch {
    return undefined;
  }
}

/** 将组件支持的各种 URI 输入统一解析为可交给解析层的 File。 */
export async function normalizeOfficeFileUri(
  uri: OfficeFileViewerUri,
  messages: OfficeFileViewerMessages,
  signal?: AbortSignal,
): Promise<NormalizedOfficeFileUri> {
  const resolvedUri = typeof uri === 'function' ? await uri(signal) : uri;

  if (resolvedUri instanceof File) return { file: resolvedUri };
  if (resolvedUri instanceof Response) {
    return {
      file: await createFileFromResponse(resolvedUri, messages),
      sourceUrl: getRemoteSourceUrl(resolvedUri.url),
    };
  }
  if (resolvedUri instanceof Blob) {
    return { file: createFileFromBlob(resolvedUri, messages) };
  }
  if (typeof resolvedUri === 'string') {
    return {
      file: await downloadOfficeFile(resolvedUri, messages, signal),
      sourceUrl: getRemoteSourceUrl(resolvedUri),
    };
  }

  throw new OfficeFileViewerInputError(messages.file.invalidUri);
}
