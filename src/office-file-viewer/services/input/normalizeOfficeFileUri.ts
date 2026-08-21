import type { OfficeFileViewerMessages } from '../../locale';
import {
  OfficeFileViewerError,
  type OfficeFileViewerErrorCode,
  type OfficeFileViewerErrorStage,
} from '../errors/OfficeFileViewerError';
import { isSupportedOfficeFileName } from '../parsing/detectPreviewKind';
import { OFFICE_MIME_EXTENSION_MAP } from '../parsing/formatDefinitions';

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

/** 从 URL 路径末段提取并解码文件名；无法解析时返回 undefined。 */
function getFileNameFromUrl(url: string): string | undefined {
  try {
    const base =
      typeof window === 'undefined' ? undefined : window.location.href;
    const parsedUrl = base ? new URL(url, base) : new URL(url);
    const lastSegment = parsedUrl.pathname.split('/').filter(Boolean).pop();
    return lastSegment ? decodeURIComponent(lastSegment) : undefined;
  } catch {
    const path = url.split(/[?#]/)[0];
    const lastSegment = path.split('/').filter(Boolean).pop();
    if (!lastSegment) return undefined;
    try {
      return decodeURIComponent(lastSegment);
    } catch {
      return lastSegment;
    }
  }
}

/** 从 Content-Disposition 响应头提取 UTF-8 或普通文件名。 */
function getFileNameFromContentDisposition(
  contentDisposition: string | null,
): string | undefined {
  if (!contentDisposition) return undefined;
  const encodedMatch = contentDisposition.match(/filename\*=UTF-8''([^;]+)/i);
  const value =
    encodedMatch?.[1] ??
    contentDisposition.match(/filename="?([^";]+)"?/i)?.[1];
  if (!value) return undefined;
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
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

/** 判断异常是否来自调用方取消，取消不应转换为加载失败。 */
function isAbortError(error: unknown) {
  return error instanceof Error && error.name === 'AbortError';
}

/** 标记由预览器输入校验产生、可安全直接展示给用户的错误。 */
export class OfficeFileViewerInputError extends OfficeFileViewerError {
  constructor(
    code: Extract<
      OfficeFileViewerErrorCode,
      | 'INVALID_INPUT'
      | 'UNSUPPORTED_FILE_TYPE'
      | 'FILE_NAME_UNRECOGNIZED'
      | 'FILE_LOAD_FAILED'
      | 'FILE_DOWNLOAD_FAILED'
    >,
    message: string,
    stage: Extract<OfficeFileViewerErrorStage, 'input' | 'download'> = 'input',
    cause?: unknown,
  ) {
    super(code, message, {
      stage,
      recoverable:
        code === 'FILE_LOAD_FAILED' || code === 'FILE_DOWNLOAD_FAILED',
      cause,
    });
    this.name = 'OfficeFileViewerInputError';
  }
}

/** 校验文件扩展名是否属于当前组件支持的 Office 格式。 */
export function ensureSupportedOfficeFile(
  file: File,
  messages: OfficeFileViewerMessages,
): void {
  if (!isSupportedOfficeFileName(file.name)) {
    throw new OfficeFileViewerInputError(
      'UNSUPPORTED_FILE_TYPE',
      messages.file.unsupported,
    );
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
    throw new OfficeFileViewerInputError(
      'FILE_NAME_UNRECOGNIZED',
      messages.file.unrecognized,
    );
  }

  const extension = getExtensionFromMimeType(blob.type);
  const inferredFileName =
    fileName && isSupportedOfficeFileName(fileName)
      ? fileName
      : extension
      ? `office-file${extension}`
      : undefined;
  if (!inferredFileName) {
    throw new OfficeFileViewerInputError(
      'FILE_NAME_UNRECOGNIZED',
      messages.file.unrecognized,
    );
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
      'FILE_DOWNLOAD_FAILED',
      messages.file.downloadFailed(response.status, response.statusText),
      'download',
    );
  }

  let blob: Blob;
  try {
    blob = await response.blob();
  } catch (error) {
    if (isAbortError(error)) throw error;
    throw new OfficeFileViewerInputError(
      'FILE_DOWNLOAD_FAILED',
      messages.file.loadFailed,
      'download',
      error,
    );
  }
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
    throw new OfficeFileViewerInputError(
      'UNSUPPORTED_FILE_TYPE',
      messages.file.unsupported,
    );
  }

  try {
    const response = await fetch(url, { signal });
    return await createFileFromResponse(response, messages, urlFileName);
  } catch (error) {
    if (isAbortError(error) || error instanceof OfficeFileViewerInputError) {
      throw error;
    }
    throw new OfficeFileViewerInputError(
      'FILE_DOWNLOAD_FAILED',
      messages.file.loadFailed,
      'download',
      error,
    );
  }
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
  let resolvedUri: File | Blob | string | Response;
  try {
    resolvedUri = typeof uri === 'function' ? await uri(signal) : uri;
  } catch (error) {
    if (isAbortError(error)) throw error;
    throw new OfficeFileViewerInputError(
      'FILE_LOAD_FAILED',
      messages.file.loadFailed,
      'input',
      error,
    );
  }

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

  throw new OfficeFileViewerInputError(
    'INVALID_INPUT',
    messages.file.invalidUri,
  );
}
