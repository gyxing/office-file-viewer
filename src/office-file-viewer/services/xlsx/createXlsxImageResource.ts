import { bytesToDataUrl, imageMimeType } from '../../shared/ooxml/media';
import {
  convertOfficeImageBlob,
  convertOfficeImageBytes,
  officeImageMimeType,
} from '../media/officeMetafile';
import type { OfficeResourceSource } from '../resource-store';
import type { XlsxPackageContext } from './XlsxPackageContext';

/** 把浏览器不能直接显示的 EMF/WMF 转换为独立 SVG。 */
export function convertXlsxImageBytes(
  path: string,
  bytes: Uint8Array,
): { bytes: Uint8Array; mimeType: string } {
  return convertOfficeImageBytes(path, bytes);
}

/** 为物化解析路径生成可直接渲染的图片 Data URL。 */
export function xlsxImageBytesToDataUrl(path: string, bytes: Uint8Array) {
  const converted = convertXlsxImageBytes(path, bytes);
  return bytesToDataUrl(converted.bytes, converted.mimeType);
}

/** 为 XLSX 建立按需图片资源，元文件只在进入可见范围时转换。 */
export function createXlsxImageResource(
  context: XlsxPackageContext,
  imagePath: string,
): OfficeResourceSource | undefined {
  if (!context.reader.has(imagePath)) return undefined;
  const entry = context.reader
    .list()
    .find((candidate) => candidate.path === imagePath);
  return {
    kind: 'lazy',
    id: [context.sessionId, 'xlsx', imagePath].join(':'),
    mimeType: officeImageMimeType(imagePath),
    size: entry?.uncompressedSize ?? 0,
    async load(signal) {
      const blob = await context.reader.readBlob(
        imagePath,
        imageMimeType(imagePath),
        signal,
      );
      return convertOfficeImageBlob(imagePath, blob);
    },
  };
}
