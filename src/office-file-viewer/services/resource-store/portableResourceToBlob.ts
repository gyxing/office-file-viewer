import type { PortableResource } from '../parsing/protocol/messages';

function rgbaToPngBlob(
  resource: Extract<PortableResource, { encoding: 'rgba' }>,
) {
  if (typeof document === 'undefined') {
    throw new Error('当前环境没有 Canvas，无法转换 DIB');
  }
  const canvas = document.createElement('canvas');
  canvas.width = resource.width;
  canvas.height = resource.height;
  const context = canvas.getContext('2d');
  if (!context) throw new Error('无法创建 DIB Canvas 上下文');
  const pixels = new Uint8ClampedArray(resource.buffer);
  const imageData = context.createImageData(resource.width, resource.height);
  imageData.data.set(pixels);
  context.putImageData(imageData, 0, 0);
  return new Promise<Blob>((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error('DIB 转 PNG 失败'));
    }, 'image/png');
  });
}

/** 把 Worker 可移植资源转换为主线程可创建 Object URL 的 Blob。 */
export async function portableResourceToBlob(resource: PortableResource) {
  if (resource.encoding === 'binary') {
    return new Blob([resource.buffer], { type: resource.mimeType });
  }
  if (resource.encoding === 'text') {
    return new Blob([resource.text], {
      type: 'image/svg+xml;charset=utf-8',
    });
  }
  return rgbaToPngBlob(resource);
}
