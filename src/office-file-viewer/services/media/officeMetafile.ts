import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import { createMediaStore, imageMimeType } from '../../shared/ooxml/media';
import { parseEmf } from '../xls/drawing/metafile/parseEmf';
import { parseWmf } from '../xls/drawing/metafile/parseWmf';
import { vectorSceneToSvg } from '../xls/drawing/metafile/vectorSceneToSvg';

/** 当前支持转换的 Office 图元文件格式。 */
type OfficeMetafileFormat = 'emf' | 'wmf';

/** 判断 Office 媒体资源是否为浏览器不能直接显示的元文件。 */
export function officeMetafileFormat(
  path: string,
): OfficeMetafileFormat | undefined {
  const lower = path.toLowerCase();
  if (lower.endsWith('.emf')) return 'emf';
  if (lower.endsWith('.wmf')) return 'wmf';
  return undefined;
}

/** 返回 Office 图片资源最终交付给浏览器的 MIME 类型。 */
export function officeImageMimeType(path: string) {
  return officeMetafileFormat(path) ? 'image/png' : imageMimeType(path);
}

function dataUrlToBlob(dataUrl: string) {
  const separator = dataUrl.indexOf(',');
  if (separator < 0) return undefined;
  const header = dataUrl.slice(0, separator);
  const payload = dataUrl.slice(separator + 1);
  const mimeType = /^data:([^;,]+)/.exec(header)?.[1] ?? 'image/png';
  const binary = header.includes(';base64')
    ? atob(payload)
    : decodeURIComponent(payload);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new Blob([bytes], { type: mimeType });
}

/** 同步转换物化解析路径中的 Office 图片字节。 */
export function convertOfficeImageBytes(path: string, bytes: Uint8Array) {
  const format = officeMetafileFormat(path);
  if (!format) return { bytes, mimeType: imageMimeType(path) };
  const scene = format === 'emf' ? parseEmf(bytes) : parseWmf(bytes);
  return {
    bytes: new TextEncoder().encode(vectorSceneToSvg(scene)),
    mimeType: 'image/svg+xml',
  };
}

/** 按需把 EMF/WMF 转为浏览器可显示图片，复杂转换失败时使用现有解析器兜底。 */
export async function convertOfficeImageBlob(path: string, blob: Blob) {
  const format = officeMetafileFormat(path);
  if (!format) return blob;
  const bytes = new Uint8Array(await blob.arrayBuffer());
  try {
    // 仅元文件动态加载完整转换器，普通 Office 图片不增加首屏下载量。
    const { convertEmfToDataUrl, convertWmfToDataUrl } = await import(
      'emf-converter'
    );
    const buffer = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer;
    const convert =
      format === 'emf' ? convertEmfToDataUrl : convertWmfToDataUrl;
    const dataUrl = await convert(buffer, { dpiScale: 1 });
    const converted = dataUrl ? dataUrlToBlob(dataUrl) : undefined;
    if (converted) return converted;
  } catch {
    // 单张复杂元文件不能中断整份文档，继续交由轻量解析器恢复可见内容。
  }
  const converted = convertOfficeImageBytes(path, bytes);
  return new Blob([converted.bytes], { type: converted.mimeType });
}
/** 收集物化 OOXML 媒体，并在登记前统一转换浏览器不支持的元文件。 */
export async function collectRenderableOfficeMedia(
  entries: OfficeEntryMap,
  mediaPrefix: string,
) {
  const media = createMediaStore();
  for (const [path, value] of entries) {
    if (!path.startsWith(mediaPrefix) || !(value instanceof Uint8Array)) {
      continue;
    }
    if (!officeMetafileFormat(path)) {
      media.register(path, value);
      continue;
    }
    // 物化模式本就会一次性读取媒体，顺序转换可限制多个复杂画布同时占用内存。
    const source = new Blob([value], { type: imageMimeType(path) });
    const converted = await convertOfficeImageBlob(path, source);
    const bytes = new Uint8Array(await converted.arrayBuffer());
    media.register(path, bytes, converted.type || 'image/png');
  }
  return media.store;
}
