import { createResourceReference } from '../parsing/assembly/resourceReferences';
import type { PortableResource } from '../parsing/protocol/messages';
import type { DocImage } from './types';

function readImageUint16(view: DataView, offset: number) {
  return view.getUint16(offset, true);
}

function readImageUint32BE(view: DataView, offset: number) {
  return view.getUint32(offset, false);
}

function readImageUint16BE(view: DataView, offset: number) {
  return view.getUint16(offset, false);
}

function readImageInt16(view: DataView, offset: number) {
  return view.getInt16(offset, true);
}

function docTwipToPx(value: number) {
  return (value / 1440) * 96;
}

/** 等待分配资源标识和地址的 DOC 图片。 */
type DocImageCandidate = Omit<DocImage, 'id' | 'src'> & {
  /** 原始字节序列。 */
  bytes: Uint8Array;
  /** 在所属数据范围中的偏移位置。 */
  offset: number;
  /** 占用或消费的字节数。 */
  byteLength: number;
  /** 候选图片是否来自文档内嵌媒体包。 */
  packagedMedia: boolean;
  /** 候选图片是否为 Web 扩展对象的预览图。 */
  webExtensionPreview: boolean;
  /** 图片资源所在的复合文档流名称。 */
  streamName: string;
};

/** 提取并汇总 `extractImageAt` 返回的数据。 */
function extractImageAt(bytes: Uint8Array, start: number) {
  if (
    bytes[start] === 0x89 &&
    bytes[start + 1] === 0x50 &&
    bytes[start + 2] === 0x4e &&
    bytes[start + 3] === 0x47 &&
    bytes[start + 4] === 0x0d &&
    bytes[start + 5] === 0x0a &&
    bytes[start + 6] === 0x1a &&
    bytes[start + 7] === 0x0a
  ) {
    let offset = start + 8;
    while (offset + 12 <= bytes.length) {
      const view = new DataView(
        bytes.buffer,
        bytes.byteOffset,
        bytes.byteLength,
      );
      const chunkLength = readImageUint32BE(view, offset);
      const chunkType = String.fromCharCode(
        bytes[offset + 4],
        bytes[offset + 5],
        bytes[offset + 6],
        bytes[offset + 7],
      );
      const nextOffset = offset + 12 + chunkLength;
      if (nextOffset > bytes.length) break;
      offset = nextOffset;
      if (chunkType === 'IEND') {
        return {
          mimeType: 'image/png',
          bytes: bytes.slice(start, offset),
        };
      }
    }
  }

  if (
    bytes[start] === 0xff &&
    bytes[start + 1] === 0xd8 &&
    bytes[start + 2] === 0xff
  ) {
    for (let index = start + 2; index + 1 < bytes.length; index += 1) {
      if (bytes[index] === 0xff && bytes[index + 1] === 0xd9) {
        return {
          mimeType: 'image/jpeg',
          bytes: bytes.slice(start, index + 2),
        };
      }
    }
  }

  return undefined;
}

function readImageSize(bytes: Uint8Array, mimeType: string) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  if (mimeType === 'image/png' && bytes.length >= 24) {
    return {
      width: readImageUint32BE(view, 16),
      height: readImageUint32BE(view, 20),
    };
  }

  if (mimeType === 'image/jpeg') {
    let offset = 2;
    while (offset + 9 < bytes.length) {
      if (bytes[offset] !== 0xff) {
        offset += 1;
        continue;
      }
      const marker = bytes[offset + 1];
      const length = readImageUint16BE(view, offset + 2);
      if (marker >= 0xc0 && marker <= 0xc3 && offset + 8 < bytes.length) {
        return {
          width: readImageUint16BE(view, offset + 7),
          height: readImageUint16BE(view, offset + 5),
        };
      }
      offset += Math.max(2, length + 2);
    }
  }

  return {};
}

/** 从 DOC 的 PICF 记录恢复随文图片经过缩放后的最终显示尺寸。 */
function readInlinePictureLayouts(bytes: Uint8Array) {
  const layouts: Array<{
    start: number;
    end: number;
    width: number;
    height: number;
  }> = [];
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  for (let offset = 0; offset + 68 <= bytes.length; offset += 1) {
    const recordLength = view.getInt32(offset, true);
    const headerLength = readImageUint16(view, offset + 4);
    const mappingMode = readImageInt16(view, offset + 6);
    if (
      recordLength < 68 ||
      offset + recordLength > bytes.length ||
      headerLength !== 0x44 ||
      (mappingMode !== 0x64 && mappingMode !== 0x66) ||
      readImageUint16(view, offset + 66) !== 0
    ) {
      continue;
    }

    const initialWidth = readImageInt16(view, offset + 28);
    const initialHeight = readImageInt16(view, offset + 30);
    const horizontalScale = readImageUint16(view, offset + 32);
    const verticalScale = readImageUint16(view, offset + 34);
    if (
      initialWidth <= 0 ||
      initialHeight <= 0 ||
      horizontalScale <= 0 ||
      verticalScale <= 0
    ) {
      continue;
    }

    layouts.push({
      start: offset,
      end: offset + recordLength,
      // PICMID 的缩放值以千分之一表示，目标尺寸以 twip 表示。
      width: docTwipToPx((initialWidth * horizontalScale) / 1000),
      height: docTwipToPx((initialHeight * verticalScale) / 1000),
    });
    offset += recordLength - 1;
  }

  return layouts;
}

function textNearBytes(
  bytes: Uint8Array,
  start: number,
  before = 320,
  after = 80,
) {
  const slice = bytes.slice(
    Math.max(0, start - before),
    Math.min(bytes.length, start + after),
  );
  return Array.from(slice, (value) =>
    value >= 32 && value <= 126 ? String.fromCharCode(value) : ' ',
  ).join('');
}

function isLikelySameImageObject(
  left: DocImageCandidate,
  right: DocImageCandidate,
) {
  if (left.streamName !== right.streamName) return false;
  if (
    left.mimeType !== right.mimeType ||
    !left.width ||
    !left.height ||
    !right.width ||
    !right.height
  )
    return false;
  if (left.width !== right.width || left.height !== right.height) return false;

  const byteDelta = Math.abs(left.byteLength - right.byteLength);
  const isCloseLength =
    byteDelta <= 1024 ||
    byteDelta / Math.max(left.byteLength, right.byteLength) <= 0.02;
  const isOfficePreviewPair =
    (left.packagedMedia && right.webExtensionPreview) ||
    (left.webExtensionPreview && right.packagedMedia);
  const isNearAlternatePreview = Math.abs(left.offset - right.offset) <= 120000;

  return isCloseLength && isOfficePreviewPair && isNearAlternatePreview;
}

function chooseBetterImageCandidate(
  left: DocImageCandidate,
  right: DocImageCandidate,
) {
  if (left.packagedMedia !== right.packagedMedia)
    return left.packagedMedia ? left : right;
  if (left.byteLength !== right.byteLength)
    return left.byteLength > right.byteLength ? left : right;
  return left.offset <= right.offset ? left : right;
}

/** 将输入标准化为 `normalizeImageCandidates` 返回的结构。 */
function normalizeImageCandidates(
  candidates: DocImageCandidate[],
  resources: PortableResource[],
) {
  const normalized: DocImageCandidate[] = [];

  candidates.forEach((candidate) => {
    const duplicateIndex = normalized.findIndex((image) =>
      isLikelySameImageObject(image, candidate),
    );
    if (duplicateIndex === -1) {
      normalized.push(candidate);
      return;
    }

    normalized[duplicateIndex] = chooseBetterImageCandidate(
      normalized[duplicateIndex],
      candidate,
    );
  });

  return normalized
    .sort((left, right) => left.offset - right.offset)
    .map(
      (
        {
          byteLength,
          packagedMedia,
          webExtensionPreview,
          streamName,
          bytes,
          ...image
        },
        index,
      ) => {
        const resourceId = `doc:image:${index + 1}`;
        const resourceBytes = bytes.slice();
        resources.push({
          id: resourceId,
          encoding: 'binary',
          mimeType: image.mimeType,
          buffer: resourceBytes.buffer,
        });
        return {
          ...image,
          id: `doc-image-${index + 1}`,
          src: createResourceReference(resourceId),
        };
      },
    );
}

/** 提取并汇总 `extractDocImagesFromStream` 返回的数据。 */
function extractDocImagesFromStream(bytes: Uint8Array, streamName: string) {
  const candidates: DocImageCandidate[] = [];
  const seen = new Set<string>();
  const pictureLayouts = readInlinePictureLayouts(bytes);
  let pictureLayoutIndex = 0;
  const signatures = [
    { mimeType: 'image/png', header: [0x89, 0x50, 0x4e, 0x47] },
    { mimeType: 'image/jpeg', header: [0xff, 0xd8, 0xff] },
  ];

  for (let index = 0; index < bytes.length - 4; index += 1) {
    const signature = signatures.find(({ header }) =>
      header.every(
        (value, headerIndex) => bytes[index + headerIndex] === value,
      ),
    );
    if (!signature) continue;

    const extracted = extractImageAt(bytes, index);
    if (!extracted || extracted.bytes.length < 128) continue;

    const head = Array.from(extracted.bytes.slice(0, 16)).join(',');
    const tail = Array.from(
      extracted.bytes.slice(Math.max(0, extracted.bytes.length - 16)),
    ).join(',');
    const key = `${extracted.mimeType}:${extracted.bytes.length}:${head}:${tail}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const context = textNearBytes(bytes, index);
    while (
      pictureLayoutIndex < pictureLayouts.length &&
      pictureLayouts[pictureLayoutIndex].end <= index
    ) {
      pictureLayoutIndex += 1;
    }
    const pictureLayout = pictureLayouts[pictureLayoutIndex];
    const displaySize =
      pictureLayout &&
      pictureLayout.start <= index &&
      pictureLayout.end >= index + extracted.bytes.length
        ? {
            width: pictureLayout.width,
            height: pictureLayout.height,
          }
        : readImageSize(extracted.bytes, extracted.mimeType);

    candidates.push({
      bytes: extracted.bytes,
      mimeType: extracted.mimeType,
      offset: index,
      byteLength: extracted.bytes.length,
      packagedMedia: /drs\/media|drs\\media/.test(context),
      webExtensionPreview: /drs\/webExtensions|drs\\webExtensions/.test(
        context,
      ),
      streamName,
      ...displaySize,
    });
  }

  return candidates;
}

/** 从 DOC 复合文档流提取、去重并注册图片资源。 */
export function extractDocImages(
  streams: Iterable<readonly [string, Uint8Array]>,
  resources: PortableResource[],
): DocImage[] {
  const candidates = Array.from(streams).flatMap(([streamName, stream]) =>
    extractDocImagesFromStream(stream, streamName),
  );

  return normalizeImageCandidates(candidates, resources);
}
