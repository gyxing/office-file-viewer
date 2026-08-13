import { convertOfficeImageBlob } from '../media/officeMetafile';
import { createResourceReference } from '../parsing/assembly/resourceReferences';
import type { PortableResource } from '../parsing/protocol/messages';
import {
  isDocObjectPreviewStreamPath,
  readDocObjectEmfPreview,
} from './docObjectPreview';
import type { DocImage } from './types';

/** 从 DOC ObjectPool 提取嵌入对象的 EMF 静态预览并注册可渲染资源。 */
export async function extractDocObjectPreviews(
  streams: Iterable<readonly [string, Uint8Array]>,
  resources: PortableResource[],
): Promise<DocImage[]> {
  const images: DocImage[] = [];

  for (const [streamPath, source] of streams) {
    if (!isDocObjectPreviewStreamPath(streamPath)) continue;
    const preview = readDocObjectEmfPreview(source);
    if (!preview) continue;

    // 完整转换器仅在确实存在 OLE/Visio 预览时加载，普通 DOC 不承担额外成本。
    const converted = await convertOfficeImageBlob(
      `${streamPath}.emf`,
      new Blob([preview.bytes], { type: 'image/x-emf' }),
    );
    const bytes = new Uint8Array(await converted.arrayBuffer());
    const imageIndex = images.length + 1;
    const resourceId = `doc:object-preview:${imageIndex}`;
    resources.push({
      id: resourceId,
      encoding: 'binary',
      mimeType: converted.type || 'image/png',
      buffer: bytes.buffer,
    });
    images.push({
      id: `doc-object-preview-${imageIndex}`,
      src: createResourceReference(resourceId),
      mimeType: converted.type || 'image/png',
      width: preview.width,
      height: preview.height,
    });
  }

  return images;
}
