import { portableResourceToBlob } from '../../resource-store/portableResourceToBlob';
import type { PortableResource } from '../protocol/messages';
import { readResourceReference } from './resourceReferences';

/** 在主线程创建和管理解析资源的 Blob URL。 */
export class ResourceRegistry {
  private readonly urls = new Map<string, string>();
  private readonly ownedUrls = new Set<string>();

  async register(resource: PortableResource): Promise<string> {
    const existing = this.urls.get(resource.id);
    if (existing) return existing;
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

  dispose() {
    if (
      typeof URL !== 'undefined' &&
      typeof URL.revokeObjectURL === 'function'
    ) {
      this.ownedUrls.forEach((url) => URL.revokeObjectURL(url));
    }
    this.ownedUrls.clear();
    this.urls.clear();
  }
}
