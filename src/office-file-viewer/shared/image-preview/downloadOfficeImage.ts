import type { OfficeImagePreviewTarget } from './types';

/** 常见图片 MIME 类型对应的下载扩展名。 */
const IMAGE_EXTENSION_BY_MIME_TYPE: Record<string, string> = {
  'image/bmp': 'bmp',
  'image/gif': 'gif',
  'image/jpeg': 'jpg',
  'image/png': 'png',
  'image/svg+xml': 'svg',
  'image/tiff': 'tiff',
  'image/webp': 'webp',
  'image/x-emf': 'emf',
  'image/x-wmf': 'wmf',
};

/** Windows 与浏览器下载名都不应包含的保留字符。 */
const INVALID_DOWNLOAD_NAME_PATTERN = /[<>:"/\\|?*\u0000-\u001f]/g;
/** 用于判断候选下载名是否已经包含常见扩展名。 */
const DOWNLOAD_NAME_EXTENSION_PATTERN = /\.[a-z0-9]{2,6}$/i;
/** 下载文件名的最大长度，避免不同系统截断后产生不可用名称。 */
const MAX_DOWNLOAD_NAME_LENGTH = 120;
/** 延迟释放临时地址，避免部分浏览器尚未接管下载就失去资源。 */
const DOWNLOAD_URL_REVOKE_DELAY_MS = 1000;

function getTargetMimeType(target: OfficeImagePreviewTarget) {
  if (target.mimeType) return target.mimeType.toLowerCase();
  return typeof target.source === 'string' || target.source.kind === 'url'
    ? undefined
    : target.source.mimeType.toLowerCase();
}

/** 根据图片元数据生成稳定且可由浏览器保存的文件名。 */
export function getOfficeImageDownloadName(target: OfficeImagePreviewTarget) {
  const candidate = (target.name || target.alt || target.id || 'office-image')
    .trim()
    .replace(INVALID_DOWNLOAD_NAME_PATTERN, '_')
    .replace(/[.\s]+$/g, '')
    .slice(0, MAX_DOWNLOAD_NAME_LENGTH);
  const safeName = candidate || 'office-image';
  if (DOWNLOAD_NAME_EXTENSION_PATTERN.test(safeName)) return safeName;
  const extension =
    IMAGE_EXTENSION_BY_MIME_TYPE[getTargetMimeType(target) ?? ''];
  return extension ? `${safeName}.${extension}` : safeName;
}

function isCrossOriginHttpUrl(url: string, ownerDocument: Document) {
  try {
    const parsed = new URL(url, ownerDocument.baseURI);
    return (
      (parsed.protocol === 'http:' || parsed.protocol === 'https:') &&
      parsed.origin !== ownerDocument.location?.origin
    );
  } catch {
    return false;
  }
}

/** 跨域图片先复制为同源临时地址，避免 download 属性被浏览器忽略。 */
async function createDownloadUrl(url: string, ownerDocument: Document) {
  if (!isCrossOriginHttpUrl(url, ownerDocument)) {
    return { url, revoke: undefined };
  }
  const view = ownerDocument.defaultView;
  const fetchResource = view?.fetch?.bind(view) ?? fetch;
  const urlApi = view?.URL ?? URL;
  const response = await fetchResource(url);
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} ${response.statusText}`);
  }
  const objectUrl = urlApi.createObjectURL(await response.blob());
  return { url: objectUrl, revoke: () => urlApi.revokeObjectURL(objectUrl) };
}

/** 通过临时链接下载图片，并在跨域资源复制完成后释放临时地址。 */
export async function downloadOfficeImage(
  url: string,
  target: OfficeImagePreviewTarget,
  ownerDocument: Document,
) {
  const downloadable = await createDownloadUrl(url, ownerDocument);
  const anchor = ownerDocument.createElement('a');
  anchor.href = downloadable.url;
  anchor.download = getOfficeImageDownloadName(target);
  anchor.rel = 'noopener noreferrer';
  anchor.style.display = 'none';
  ownerDocument.body.appendChild(anchor);
  try {
    anchor.click();
  } finally {
    anchor.remove();
    if (downloadable.revoke) {
      const view = ownerDocument.defaultView;
      const scheduleCleanup = view ? view.setTimeout.bind(view) : setTimeout;
      scheduleCleanup(downloadable.revoke, DOWNLOAD_URL_REVOKE_DELAY_MS);
    }
  }
}
