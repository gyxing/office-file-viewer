import type { OfficeEntryMap } from './archive';
import { readBinary } from './archive';

/** 按包内路径和文件名索引的 OOXML 媒体地址。 */
export type MediaStore = {
  /** 按 OOXML 包内完整路径索引的媒体地址。 */
  byPath: Record<string, string>;
  /** 按媒体文件名索引的媒体地址。 */
  byName: Record<string, string>;
};

/** OOXML 关系标识、目标路径和内容类型。 */
export type OfficeRelationship = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 关系指向的包内路径或外部地址。 */
  target: string;
  /** 关系定义的目标内容类型。 */
  type?: string;
};

/** 按关系文件路径和关系标识组织的 OOXML 关系。 */
export type OfficeRelationshipMap = Record<
  string,
  Record<string, OfficeRelationship>
>;

function bytesToBase64(bytes: Uint8Array) {
  let binary = '';
  for (let index = 0; index < bytes.length; index += 1) {
    binary += String.fromCharCode(bytes[index]);
  }
  return btoa(binary);
}

/** 将二进制资源转换为可直接使用的数据地址。 */
export function bytesToDataUrl(bytes: Uint8Array, contentType = 'image/png') {
  return `data:${contentType};base64,${bytesToBase64(bytes)}`;
}

/** 根据文件扩展名返回浏览器可识别的图片 MIME 类型。 */
export function imageMimeType(path: string) {
  const lower = path.toLowerCase();
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
  if (lower.endsWith('.gif')) return 'image/gif';
  if (lower.endsWith('.svg')) return 'image/svg+xml';
  if (lower.endsWith('.webp')) return 'image/webp';
  if (lower.endsWith('.bmp') || lower.endsWith('.dib')) return 'image/bmp';
  if (lower.endsWith('.emf')) return 'image/x-emf';
  if (lower.endsWith('.wmf')) return 'image/x-wmf';
  return 'image/png';
}

/** 创建可按包内路径和文件名查询的媒体存储。 */
export function createMediaStore() {
  const store: MediaStore = {
    byPath: {},
    byName: {},
  };

  function register(
    path: string,
    bytes: Uint8Array,
    contentType = imageMimeType(path),
  ) {
    const dataUrl = bytesToDataUrl(bytes, contentType);
    const fileName = path.split('/').pop() ?? path;
    store.byPath[path] = dataUrl;
    store.byName[fileName] = dataUrl;
    return dataUrl;
  }

  function resolve(pathOrName?: string) {
    if (!pathOrName) {
      return undefined;
    }
    return store.byPath[pathOrName] ?? store.byName[pathOrName];
  }

  return { store, register, resolve };
}

/** 将输入标准化为 `normalizeRelationshipTarget` 返回的结构。 */
export function normalizeRelationshipTarget(relsPath: string, target: string) {
  if (/^[a-z]+:/i.test(target)) {
    return target;
  }

  const baseDir = relsPath
    .replace(/\/_rels\/[^/]+\.rels$/, '')
    .replace(/\/[^/]+\.rels$/, '');
  const parts = `${baseDir}/${target}`.split('/');
  const normalized: string[] = [];
  parts.forEach((part) => {
    if (!part || part === '.') return;
    if (part === '..') {
      normalized.pop();
      return;
    }
    normalized.push(part);
  });
  return normalized.join('/');
}

/** 解析并确定 `resolvePackageMediaRef` 对应的引用或配置。 */
export function resolvePackageMediaRef<T>(
  target: string | undefined,
  mediaByPath: Record<string, T>,
  mediaByName: Record<string, T>,
  rootDir: string,
) {
  if (!target) return undefined;
  const fileName = target.split('/').pop() ?? target;
  return (
    mediaByPath[target] ??
    mediaByPath[
      `${rootDir}/${target.replace(new RegExp(`^${rootDir}/`), '')}`
    ] ??
    mediaByName[fileName]
  );
}

/** 提取并汇总 `collectMedia` 返回的数据。 */
export function collectMedia(entries: OfficeEntryMap, mediaPrefix: string) {
  const media = createMediaStore();

  for (const [path] of entries) {
    if (!path.startsWith(mediaPrefix)) continue;
    const binary = readBinary(entries, path);
    if (!binary) continue;
    media.register(path, binary);
  }

  return media.store;
}
