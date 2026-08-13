/** DOC ObjectPool 中嵌入对象的打印预览流名称。 */
const DOC_OBJECT_PREVIEW_PATH = /^\/?ObjectPool\/[^/]+\/\u0003EPRINT$/i;

/** EMF 文件头固定签名。 */
const EMF_SIGNATURE = 0x464d4520;
/** EMF 文件头的最小字节数。 */
const EMF_HEADER_BYTES = 88;
/** Word 对象预览可能附带短前缀，只在有限范围内寻找真实 EMF 头。 */
const MAX_EMF_PREFIX_BYTES = 512;

/** 从 DOC 嵌入对象预览中读取出的 EMF 资源及其自然尺寸。 */
export type DocObjectEmfPreview = {
  /** 去除可选包装前缀后的完整 EMF 字节。 */
  bytes: Uint8Array;
  /** EMF 画框换算后的自然宽度，单位为标准化渲染像素。 */
  width?: number;
  /** EMF 画框换算后的自然高度，单位为标准化渲染像素。 */
  height?: number;
};

/** 判断 CFB 路径是否为 DOC 嵌入对象的静态打印预览。 */
export function isDocObjectPreviewStreamPath(path: string) {
  return DOC_OBJECT_PREVIEW_PATH.test(path.replace(/\\/g, '/'));
}

/** 将 EMF 使用的 0.01 毫米画框单位换算为 CSS 像素。 */
function emfFrameToPx(value: number) {
  return (value / 2540) * 96;
}

/** 校验并截取 ObjectPool EPRINT 流中的真实 EMF 数据。 */
export function readDocObjectEmfPreview(
  source: Uint8Array,
): DocObjectEmfPreview | undefined {
  if (source.length < EMF_HEADER_BYTES) return undefined;
  const view = new DataView(
    source.buffer,
    source.byteOffset,
    source.byteLength,
  );
  const searchEnd = Math.min(
    source.length - EMF_HEADER_BYTES,
    MAX_EMF_PREFIX_BYTES,
  );

  for (let offset = 0; offset <= searchEnd; offset += 1) {
    if (
      view.getUint32(offset, true) !== 1 ||
      view.getUint32(offset + 40, true) !== EMF_SIGNATURE
    ) {
      continue;
    }
    const byteLength = view.getUint32(offset + 48, true);
    if (byteLength < EMF_HEADER_BYTES || offset + byteLength > source.length) {
      continue;
    }
    const left = view.getInt32(offset + 24, true);
    const top = view.getInt32(offset + 28, true);
    const right = view.getInt32(offset + 32, true);
    const bottom = view.getInt32(offset + 36, true);
    const frameWidth = right - left;
    const frameHeight = bottom - top;
    return {
      bytes: source.slice(offset, offset + byteLength),
      width: frameWidth > 0 ? emfFrameToPx(frameWidth) : undefined,
      height: frameHeight > 0 ? emfFrameToPx(frameHeight) : undefined,
    };
  }

  return undefined;
}
