import {
  OFFICE_ART_RECORD,
  parseOfficeArtRecords,
  type OfficeArtRecord,
} from '../../../shared/officeart';
import type {
  PresentationImagePreviewMetadata,
  PresentationImageSourceKind,
} from '../../presentation/types';
import {
  createPptResourceId,
  registerPptResource,
  type PptParseContext,
} from '../types';
import { createPptStaticPreviewCard } from './createStaticPreviewCard';

/** PPT Pictures 流中位图的 MIME 类型和数据偏移。 */
type RasterInfo = {
  /** 资源的 MIME 类型，用于选择解码和渲染方式。 */
  mimeType: string;
  /** 原始字节序列。 */
  bytes: Uint8Array;
};

/** 查找 `findSignature` 对应的目标数据。 */
function findSignature(bytes: Uint8Array, signature: number[]) {
  const limit = Math.min(bytes.length - signature.length, 96);
  for (let offset = 0; offset <= limit; offset += 1) {
    if (signature.every((value, index) => bytes[offset + index] === value)) {
      return offset;
    }
  }
  return -1;
}

function readRaster(record: OfficeArtRecord): RasterInfo | undefined {
  const candidates = [
    {
      type: OFFICE_ART_RECORD.BLIP_PNG,
      mimeType: 'image/png',
      signature: [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a],
    },
    {
      type: OFFICE_ART_RECORD.BLIP_JPEG,
      mimeType: 'image/jpeg',
      signature: [0xff, 0xd8, 0xff],
    },
  ];
  const candidate = candidates.find((item) => item.type === record.type);
  if (!candidate) return undefined;
  const offset = findSignature(record.data, candidate.signature);
  if (offset < 0) return undefined;
  return {
    mimeType: candidate.mimeType,
    bytes: record.data.subarray(offset),
  };
}

/** 判断 OfficeArt 图片记录是否属于矢量或元文件资源。 */
function isVectorBlip(record: OfficeArtRecord) {
  return (
    record.type === OFFICE_ART_RECORD.BLIP_WMF ||
    record.type === OFFICE_ART_RECORD.BLIP_EMF ||
    record.type === OFFICE_ART_RECORD.BLIP_PICT
  );
}

/** 为二进制 PPT 图片建立不影响渲染的预览识别元数据。 */
function createPreviewMetadata(
  record: OfficeArtRecord,
  raster: RasterInfo | undefined,
  resourceKey: string,
): PresentationImagePreviewMetadata {
  const sourceKind: PresentationImageSourceKind = raster
    ? 'raster'
    : isVectorBlip(record)
    ? 'vector'
    : 'unknown';
  return {
    sourceKind,
    mimeType: raster?.mimeType,
    resourceKey,
    resourceSize: raster?.bytes.byteLength ?? record.data.byteLength,
  };
}

function fallbackLabel(type: number) {
  if (type === OFFICE_ART_RECORD.BLIP_WMF)
    return ['WMF 图像', '矢量图静态预览'];
  if (type === OFFICE_ART_RECORD.BLIP_EMF)
    return ['EMF 图像', '矢量图静态预览'];
  if (type === OFFICE_ART_RECORD.BLIP_PICT)
    return ['PICT 图像', '嵌入图像静态预览'];
  if (type === OFFICE_ART_RECORD.BLIP_DIB) return ['DIB 图像', '位图静态预览'];
  return ['嵌入图像', 'PowerPoint 97–2003 图片对象'];
}

function toExactArrayBuffer(bytes: Uint8Array) {
  return Uint8Array.from(bytes).buffer;
}

/** 将已解析的 Pictures 记录追加到一基序号资源映射。 */
export async function registerPptPictureRecords(
  records: readonly OfficeArtRecord[],
  context: PptParseContext,
) {
  const startIndex = context.blipUrls.size;
  for (let index = 0; index < records.length; index += 1) {
    const record = records[index];
    const raster = readRaster(record);
    const blipIndex = startIndex + index + 1;
    const resourceKey = `ppt-blip:${blipIndex}`;
    let reference: string;
    if (raster) {
      reference = registerPptResource(context, {
        id: createPptResourceId(context, 'picture'),
        encoding: 'binary',
        mimeType: raster.mimeType,
        buffer: toExactArrayBuffer(raster.bytes),
      });
    } else {
      const [title, detail] = fallbackLabel(record.type);
      reference = createPptStaticPreviewCard(title, detail, context);
    }
    context.blipUrls.set(blipIndex, reference);
    context.blipPreviewMetadata.set(
      blipIndex,
      createPreviewMetadata(record, raster, resourceKey),
    );
    await context.yieldIfNeeded();
  }
  return context.blipUrls;
}

/** 解析 Pictures 流并建立一基序号到可传输资源引用的映射。 */
export async function readPptPictures(
  picturesStream: Uint8Array | undefined,
  context: PptParseContext,
) {
  if (!picturesStream?.length) return context.blipUrls;
  return registerPptPictureRecords(
    parseOfficeArtRecords(picturesStream, context.warnings),
    context,
  );
}
