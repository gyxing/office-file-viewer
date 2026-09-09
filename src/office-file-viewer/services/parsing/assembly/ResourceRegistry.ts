import { portableResourceToBlob } from '../../resource-store/portableResourceToBlob';
import type { OfficeResourceDescriptor } from '../../resource-store/types';
import type { PortableResource } from '../protocol/messages';
import { readResourceReference } from './resourceReferences';

function classifyResourceKind(
  mimeType: string,
): OfficeResourceDescriptor['kind'] {
  const normalized = mimeType.toLowerCase();
  if (normalized.startsWith('image/')) return 'image';
  if (
    normalized.startsWith('audio/') ||
    normalized.startsWith('video/') ||
    normalized.includes('media')
  ) {
    return 'media';
  }
  if (
    normalized.startsWith('font/') ||
    /(?:ttf|otf|woff2?|eot)(?:$|[+;])/i.test(normalized)
  ) {
    return 'font';
  }
  return 'other';
}

function getResourceSize(resource: PortableResource) {
  if ('buffer' in resource) return resource.buffer.byteLength;
  if ('text' in resource) {
    try {
      return new Blob([resource.text]).size;
    } catch {
      return resource.text.length;
    }
  }
  return undefined;
}

function describeResource(
  resource: PortableResource,
): OfficeResourceDescriptor {
  return Object.freeze({
    id: resource.id,
    kind: classifyResourceKind(resource.mimeType),
    mimeType: resource.mimeType,
    size: getResourceSize(resource),
  });
}

/** 在主线程创建和管理解析资源的 Blob URL。 */
export class ResourceRegistry {
  private readonly urls = new Map<string, string>();
  private readonly ownedUrls = new Set<string>();
  private readonly descriptors = new Map<string, OfficeResourceDescriptor>();

  async register(resource: PortableResource): Promise<string> {
    const existing = this.urls.get(resource.id);
    if (existing) return existing;
    this.descriptors.set(resource.id, describeResource(resource));
    if (
      typeof URL === 'undefined' ||
      typeof URL.createObjectURL !== 'function'
    ) {
      throw new Error('当前环境不支持 Blob URL');
    }
    const url = URL.createObjectURL(await portableResourceToBlob(resource));
    this.urls.set(resource.id, url);
    this.ownedUrls.add(url);
    return url;
  }

  resolve(reference: string): string {
    const resourceId = readResourceReference(reference);
    if (!resourceId) return reference;
    const url = this.urls.get(resourceId);
    if (!url) {
      const error = new Error(`解析资源不存在：${resourceId}`) as Error & {
        /** ResourceRegistry 的稳定代码，用于程序化识别具体情况。 */
        code: string;
      };
      error.code = 'RESOURCE_NOT_FOUND';
      throw error;
    }
    return url;
  }

  /** 将 URL 的释放责任移交给最终文档。 */
  takeObjectUrls(): string[] {
    const urls = [...this.ownedUrls];
    this.ownedUrls.clear();
    return urls;
  }

  /** 返回当前解析会话收集的可序列化资源元数据。 */
  getResourceRefs(): readonly OfficeResourceDescriptor[] {
    return Object.freeze([...this.descriptors.values()]);
  }

  dispose() {
    if (
      typeof URL !== 'undefined' &&
      typeof URL.revokeObjectURL === 'function'
    ) {
      this.ownedUrls.forEach((url) => URL.revokeObjectURL(url));
    }
    this.ownedUrls.clear();
    this.urls.clear();
    this.descriptors.clear();
  }
}
