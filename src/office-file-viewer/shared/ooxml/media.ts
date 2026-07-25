import type { OfficeEntryMap } from './archive';
import { readBinary } from './archive';

/** 描述 MediaStore 在 OOXML 公共解析中的数据结构。 */
export type MediaStore = {
  /** MediaStore 按包内完整路径索引的媒体资源映射。 */
  byPath: Record<string, string>;
  /** MediaStore 按文件名索引的媒体资源映射。 */
  byName: Record<string, string>;
};

/** 描述 OfficeRelationship 在 OOXML 公共解析中的数据结构。 */
export type OfficeRelationship = {
  /** OfficeRelationship 在所属文档或任务中的唯一标识。 */
  id: string;
  /** OfficeRelationship 的 target 文本值。 */
  target: string;
  /** 用于区分 OfficeRelationship 不同结构分支的类型标识。 */
  type?: string;
};

/** 描述 OfficeRelationshipMap 在 OOXML 公共解析中的数据结构。 */
export type OfficeRelationshipMap = Record<
  string,
  Record<string, OfficeRelationship>
>;

/** 执行 `bytesToBase64` 封装的 OOXML 公共解析处理步骤。 */
function bytesToBase64(bytes: Uint8Array) {
  let binary = '';
  for (let index = 0; index < bytes.length; index += 1) {
    binary += String.fromCharCode(bytes[index]);
  }
  return btoa(binary);
}

/** 执行 `bytesToDataUrl` 封装的 OOXML 公共解析处理步骤。 */
export function bytesToDataUrl(bytes: Uint8Array, contentType = 'image/png') {
  return `data:${contentType};base64,${bytesToBase64(bytes)}`;
}

/** 执行 `imageMimeType` 封装的 OOXML 公共解析处理步骤。 */
export function imageMimeType(path: string) {
  const lower = path.toLowerCase();
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
  if (lower.endsWith('.gif')) return 'image/gif';
  if (lower.endsWith('.svg')) return 'image/svg+xml';
  if (lower.endsWith('.webp')) return 'image/webp';
  return 'image/png';
}

/** 创建 `createMediaStore` 返回的对象，供 OOXML 公共解析使用。 */
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
export function resolvePackageMediaRef(
  target: string | undefined,
  mediaByPath: Record<string, string>,
  mediaByName: Record<string, string>,
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
