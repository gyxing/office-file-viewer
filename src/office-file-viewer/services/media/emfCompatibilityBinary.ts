/** EMF 文件头记录。 */
export const EMR_HEADER = 1;
/** EMF 文件结束记录。 */
export const EMR_EOF = 14;
/** 32 位多折线记录。 */
export const EMR_POLYPOLYLINE = 7;
/** 16 位多折线记录。 */
export const EMR_POLYPOLYLINE16 = 90;
/** 创建普通画刷记录。 */
export const EMR_CREATEBRUSHINDIRECT = 39;
/** 创建 DIB 图案画刷记录。 */
export const EMR_CREATEDIBPATTERNBRUSHPT = 94;
/** 位块传输记录。 */
export const EMR_BITBLT = 76;
/** 拉伸 DIB 位图记录。 */
export const EMR_STRETCHDIBITS = 81;
/** 透明度混合记录。 */
export const EMR_ALPHABLEND = 114;

/** EMF 文件头中的签名字节。 */
const EMF_SIGNATURE = 0x464d4520;
/** 与转换器画布上限一致的兼容位图边长。 */
const MAX_BITMAP_DIMENSION = 8192;
/** 限制兼容位图的像素总量，避免异常记录造成瞬时大内存分配。 */
const MAX_BITMAP_PIXELS = 4 * 1024 * 1024;
/** 限制单个 EMF 的记录总量，避免异常数据造成长时间扫描。 */
const MAX_EMF_RECORDS = 200_000;
/** DPa：将目标像素与当前图案画刷做按位与。 */
const ROP_DPA = 0x00a000c9;
/** SRCCOPY：由兼容层生成的位图直接覆盖目标区域。 */
const ROP_SRCCOPY = 0x00cc0020;

/** 通过长度校验后的 EMF 记录位置。 */
export type EmfRecord = {
  /** EMR 记录类型。 */
  type: number;
  /** 记录相对 EMF 首字节的偏移。 */
  offset: number;
  /** 包含记录头在内的字节数。 */
  size: number;
};

/** 已解码的单色 DIB 图案画刷。 */
export type EmfPatternBrush = {
  /** EMF 对象表中的画刷句柄。 */
  handle: number;
  /** 创建该画刷的记录偏移。 */
  creationOffset: number;
  /** 图案宽度，单位为像素。 */
  width: number;
  /** 图案高度，单位为像素。 */
  height: number;
  /** 按从上到下顺序保存的黑色像素标记。 */
  blackPixels: boolean[];
};

/** 图案蒙版转换时需要的画刷状态。 */
export type PatternMaskContext = {
  /** 当前选中的单色图案画刷。 */
  brush: EmfPatternBrush;
  /** 当前画刷原点横坐标。 */
  originX: number;
  /** 当前画刷原点纵坐标。 */
  originY: number;
};

function createView(bytes: Uint8Array) {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

function isRangeValid(start: number, length: number, end: number) {
  return start >= 0 && length >= 0 && start <= end - length;
}

function isBitmapSizeSafe(width: number, height: number) {
  return (
    width > 0 &&
    height > 0 &&
    width <= MAX_BITMAP_DIMENSION &&
    height <= MAX_BITMAP_DIMENSION &&
    width * height <= MAX_BITMAP_PIXELS
  );
}

function isCompatibilityRecord(type: number) {
  return (
    type === EMR_POLYPOLYLINE16 ||
    type === EMR_CREATEDIBPATTERNBRUSHPT ||
    type === EMR_ALPHABLEND
  );
}

/** 快速判断是否包含兼容层关注的记录，不为普通 EMF 创建记录索引。 */
export function hasEmfCompatibilityRecords(bytes: Uint8Array) {
  if (bytes.length < 88) return false;
  const view = createView(bytes);
  if (
    view.getUint32(0, true) !== EMR_HEADER ||
    view.getUint32(40, true) !== EMF_SIGNATURE
  ) {
    return false;
  }

  let offset = 0;
  let recordCount = 0;
  while (offset + 8 <= bytes.length && recordCount < MAX_EMF_RECORDS) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (
      size < 8 ||
      size % 4 !== 0 ||
      !isRangeValid(offset, size, bytes.length)
    ) {
      return false;
    }
    if (isCompatibilityRecord(type)) return true;
    offset += size;
    recordCount += 1;
    if (type === EMR_EOF) return false;
  }
  return false;
}

/** 扫描并校验完整 EMF 记录流；结构不可信时返回 undefined。 */
export function scanEmfRecords(bytes: Uint8Array) {
  if (bytes.length < 88) return undefined;
  const view = createView(bytes);
  if (
    view.getUint32(0, true) !== EMR_HEADER ||
    view.getUint32(40, true) !== EMF_SIGNATURE
  ) {
    return undefined;
  }

  const records: EmfRecord[] = [];
  let offset = 0;
  while (offset + 8 <= bytes.length && records.length < MAX_EMF_RECORDS) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (
      size < 8 ||
      size % 4 !== 0 ||
      !isRangeValid(offset, size, bytes.length)
    ) {
      return undefined;
    }
    records.push({ type, offset, size });
    offset += size;
    if (type === EMR_EOF) {
      return offset === bytes.length ? records : undefined;
    }
  }
  return undefined;
}

/** 使用替换记录重建 EMF，并同步文件头中的总字节数与记录数。 */
export function rebuildEmf(
  bytes: Uint8Array,
  records: EmfRecord[],
  replacements: Map<number, Uint8Array>,
) {
  if (!replacements.size) return bytes;
  let outputSize = 0;
  for (const record of records) {
    outputSize += replacements.get(record.offset)?.length ?? record.size;
  }
  if (!Number.isSafeInteger(outputSize) || outputSize < 88) return bytes;

  const output = new Uint8Array(outputSize);
  let outputOffset = 0;
  for (const record of records) {
    const replacement = replacements.get(record.offset);
    const value =
      replacement ?? bytes.subarray(record.offset, record.offset + record.size);
    output.set(value, outputOffset);
    outputOffset += value.length;
  }
  const view = createView(output);
  view.setUint32(48, output.length, true);
  view.setUint32(52, records.length, true);
  return output;
}

function copyBounds(source: Uint8Array, record: EmfRecord, target: Uint8Array) {
  target.set(source.subarray(record.offset + 8, record.offset + 24), 8);
}

function createBitmapInfoHeader(
  target: DataView,
  offset: number,
  width: number,
  height: number,
  imageSize: number,
) {
  target.setUint32(offset, 40, true);
  target.setInt32(offset + 4, width, true);
  // 负高度明确指定从上到下的扫描线，避免再次翻转兼容层生成的像素。
  target.setInt32(offset + 8, -height, true);
  target.setUint16(offset + 12, 1, true);
  target.setUint16(offset + 14, 32, true);
  target.setUint32(offset + 16, 0, true);
  target.setUint32(offset + 20, imageSize, true);
}

function createStretchDibitsRecord(
  source: Uint8Array,
  record: EmfRecord,
  destination: { x: number; y: number; width: number; height: number },
  bitmap: { width: number; height: number; pixels: Uint8Array },
) {
  const fixedSize = 80;
  const bitmapHeaderSize = 40;
  const output = new Uint8Array(
    fixedSize + bitmapHeaderSize + bitmap.pixels.length,
  );
  const view = createView(output);
  view.setUint32(0, EMR_STRETCHDIBITS, true);
  view.setUint32(4, output.length, true);
  copyBounds(source, record, output);
  view.setInt32(24, destination.x, true);
  view.setInt32(28, destination.y, true);
  view.setInt32(32, 0, true);
  view.setInt32(36, 0, true);
  view.setInt32(40, bitmap.width, true);
  view.setInt32(44, bitmap.height, true);
  view.setUint32(48, fixedSize, true);
  view.setUint32(52, bitmapHeaderSize, true);
  view.setUint32(56, fixedSize + bitmapHeaderSize, true);
  view.setUint32(60, bitmap.pixels.length, true);
  view.setUint32(64, 0, true);
  view.setUint32(68, ROP_SRCCOPY, true);
  view.setInt32(72, destination.width, true);
  view.setInt32(76, destination.height, true);
  createBitmapInfoHeader(
    view,
    fixedSize,
    bitmap.width,
    bitmap.height,
    bitmap.pixels.length,
  );
  output.set(bitmap.pixels, fixedSize + bitmapHeaderSize);
  return output;
}

/** 将转换器暂不支持的 16 位多折线无损扩展为 32 位记录。 */
export function convertPolyPolyline16(bytes: Uint8Array, record: EmfRecord) {
  if (record.size < 32) return undefined;
  const view = createView(bytes);
  const polylineCount = view.getUint32(record.offset + 24, true);
  const pointCount = view.getUint32(record.offset + 28, true);
  const countsSize = polylineCount * 4;
  const pointsOffset = record.offset + 32 + countsSize;
  if (
    polylineCount > 10_000 ||
    pointCount > 100_000 ||
    !isRangeValid(pointsOffset, pointCount * 4, record.offset + record.size)
  ) {
    return undefined;
  }
  let countedPoints = 0;
  for (let index = 0; index < polylineCount; index += 1) {
    countedPoints += view.getUint32(record.offset + 32 + index * 4, true);
  }
  if (countedPoints !== pointCount) return undefined;

  const output = new Uint8Array(32 + countsSize + pointCount * 8);
  const outputView = createView(output);
  output.set(bytes.subarray(record.offset, record.offset + 32 + countsSize));
  outputView.setUint32(0, EMR_POLYPOLYLINE, true);
  outputView.setUint32(4, output.length, true);
  for (let index = 0; index < pointCount; index += 1) {
    const sourceOffset = pointsOffset + index * 4;
    const targetOffset = 32 + countsSize + index * 8;
    outputView.setInt32(targetOffset, view.getInt16(sourceOffset, true), true);
    outputView.setInt32(
      targetOffset + 4,
      view.getInt16(sourceOffset + 2, true),
      true,
    );
  }
  return output;
}

/** 兼容层读取位图所需的 DIB 元数据。 */
type DibDescriptor = {
  /** DIB 像素宽度。 */
  width: number;
  /** DIB 像素高度。 */
  height: number;
  /** 正数高度表示扫描线按从下到上保存。 */
  bottomUp: boolean;
  /** 每行像素占用的对齐字节数。 */
  stride: number;
  /** DIB 头相对 EMF 首字节的位置。 */
  headerOffset: number;
  /** 像素数据相对 EMF 首字节的位置。 */
  bitsOffset: number;
  /** 每个像素使用的位数。 */
  bitCount: number;
  /** DIB 压缩模式。 */
  compression: number;
  /** DIB 头和颜色表的总字节数。 */
  headerBytes: number;
  /** 像素数据的总字节数。 */
  bitsBytes: number;
};

function readDibDescriptor(
  bytes: Uint8Array,
  record: EmfRecord,
  headerOffset: number,
  headerBytes: number,
  bitsOffset: number,
  bitsBytes: number,
) {
  const recordEnd = record.offset + record.size;
  const headerStart = record.offset + headerOffset;
  const bitsStart = record.offset + bitsOffset;
  if (
    headerBytes < 40 ||
    !isRangeValid(headerStart, headerBytes, recordEnd) ||
    !isRangeValid(bitsStart, bitsBytes, recordEnd)
  ) {
    return undefined;
  }
  const view = createView(bytes);
  const dibHeaderSize = view.getUint32(headerStart, true);
  const width = view.getInt32(headerStart + 4, true);
  const rawHeight = view.getInt32(headerStart + 8, true);
  const planes = view.getUint16(headerStart + 12, true);
  const bitCount = view.getUint16(headerStart + 14, true);
  const compression = view.getUint32(headerStart + 16, true);
  if (
    dibHeaderSize < 40 ||
    dibHeaderSize > headerBytes ||
    width <= 0 ||
    rawHeight === 0 ||
    planes !== 1
  ) {
    return undefined;
  }
  const height = Math.abs(rawHeight);
  const stride = Math.ceil((width * bitCount) / 32) * 4;
  if (!isBitmapSizeSafe(width, height) || stride * height > bitsBytes) {
    return undefined;
  }
  const descriptor: DibDescriptor = {
    width,
    height,
    bottomUp: rawHeight > 0,
    stride,
    headerOffset: headerStart,
    bitsOffset: bitsStart,
    bitCount,
    compression,
    headerBytes,
    bitsBytes,
  };
  return descriptor;
}

function sourcePixelOffset(dib: DibDescriptor, x: number, y: number) {
  const row = dib.bottomUp ? dib.height - 1 - y : y;
  return dib.bitsOffset + row * dib.stride + x * 4;
}

function hasStandard32BitColorMasks(bytes: Uint8Array, dib: DibDescriptor) {
  if (dib.compression === 0) return true;
  if (dib.compression !== 3 || dib.headerBytes < 52) return false;
  const view = createView(bytes);
  const masksOffset = dib.headerOffset + 40;
  return (
    view.getUint32(masksOffset, true) === 0x00ff0000 &&
    view.getUint32(masksOffset + 4, true) === 0x0000ff00 &&
    view.getUint32(masksOffset + 8, true) === 0x000000ff
  );
}

function unpremultiplyColor(value: number, alpha: number) {
  if (alpha <= 0) return 0;
  if (alpha >= 255) return value;
  return Math.min(255, Math.round((value * 255) / alpha));
}

function readAlphaBlendBitmap(
  bytes: Uint8Array,
  dib: DibDescriptor,
  source: { x: number; y: number; width: number; height: number },
  constantAlpha: number,
  usesPixelAlpha: boolean,
) {
  if (
    dib.bitCount !== 32 ||
    !hasStandard32BitColorMasks(bytes, dib) ||
    source.x < 0 ||
    source.y < 0 ||
    source.x + source.width > dib.width ||
    source.y + source.height > dib.height ||
    !isBitmapSizeSafe(source.width, source.height)
  ) {
    return undefined;
  }
  const pixels = new Uint8Array(source.width * source.height * 4);
  for (let y = 0; y < source.height; y += 1) {
    for (let x = 0; x < source.width; x += 1) {
      const input = sourcePixelOffset(dib, source.x + x, source.y + y);
      const output = (y * source.width + x) * 4;
      const sourceAlpha = usesPixelAlpha ? bytes[input + 3] : 255;
      const alpha = Math.round((sourceAlpha * constantAlpha) / 255);
      if (alpha === 0) {
        // emf-converter 会把 0 误判成“无 Alpha 通道”，使用 1 保持视觉透明。
        pixels[output + 3] = 1;
        continue;
      }
      const blue = bytes[input];
      const green = bytes[input + 1];
      const red = bytes[input + 2];
      pixels[output] = usesPixelAlpha
        ? unpremultiplyColor(blue, sourceAlpha)
        : blue;
      pixels[output + 1] = usesPixelAlpha
        ? unpremultiplyColor(green, sourceAlpha)
        : green;
      pixels[output + 2] = usesPixelAlpha
        ? unpremultiplyColor(red, sourceAlpha)
        : red;
      pixels[output + 3] = Math.max(1, alpha);
    }
  }
  return pixels;
}

/** 将标准 AlphaBlend 转换为转换器可处理的透明 32 位 DIB。 */
export function convertAlphaBlend(bytes: Uint8Array, record: EmfRecord) {
  if (record.size < 108) return undefined;
  const view = createView(bytes);
  const offset = record.offset;
  const blendOperation = bytes[offset + 40];
  const constantAlpha = bytes[offset + 42];
  const alphaFormat = bytes[offset + 43];
  const destination = {
    x: view.getInt32(offset + 24, true),
    y: view.getInt32(offset + 28, true),
    width: view.getInt32(offset + 32, true),
    height: view.getInt32(offset + 36, true),
  };
  const source = {
    x: view.getInt32(offset + 44, true),
    y: view.getInt32(offset + 48, true),
    width: view.getInt32(offset + 100, true),
    height: view.getInt32(offset + 104, true),
  };
  if (
    blendOperation !== 0 ||
    (alphaFormat !== 0 && alphaFormat !== 1) ||
    !isBitmapSizeSafe(destination.width, destination.height) ||
    !isBitmapSizeSafe(source.width, source.height) ||
    view.getUint32(offset + 80, true) !== 0
  ) {
    return undefined;
  }
  const dib = readDibDescriptor(
    bytes,
    record,
    view.getUint32(offset + 84, true),
    view.getUint32(offset + 88, true),
    view.getUint32(offset + 92, true),
    view.getUint32(offset + 96, true),
  );
  if (!dib) return undefined;
  const pixels = readAlphaBlendBitmap(
    bytes,
    dib,
    source,
    constantAlpha,
    alphaFormat === 1,
  );
  if (!pixels) return undefined;
  return createStretchDibitsRecord(bytes, record, destination, {
    width: source.width,
    height: source.height,
    pixels,
  });
}

function isBlack(color: [number, number, number]) {
  return color[0] === 0 && color[1] === 0 && color[2] === 0;
}

function isWhite(color: [number, number, number]) {
  return color[0] === 255 && color[1] === 255 && color[2] === 255;
}

function readPaletteColor(
  bytes: Uint8Array,
  offset: number,
): [number, number, number] {
  return [bytes[offset + 2], bytes[offset + 1], bytes[offset]];
}

/** 解码能够用透明蒙版表达的黑白 DIB 图案画刷。 */
export function readPatternBrush(bytes: Uint8Array, record: EmfRecord) {
  if (record.size < 36) return undefined;
  const view = createView(bytes);
  const offset = record.offset;
  if (view.getUint32(offset + 12, true) !== 0) return undefined;
  const dib = readDibDescriptor(
    bytes,
    record,
    view.getUint32(offset + 16, true),
    view.getUint32(offset + 20, true),
    view.getUint32(offset + 24, true),
    view.getUint32(offset + 28, true),
  );
  if (!dib || dib.bitCount !== 1 || dib.compression !== 0) return undefined;
  const viewColors = view.getUint32(dib.headerOffset + 32, true) || 2;
  const paletteOffset =
    dib.headerOffset + view.getUint32(dib.headerOffset, true);
  if (
    viewColors < 2 ||
    !isRangeValid(paletteOffset, 8, dib.headerOffset + dib.headerBytes)
  ) {
    return undefined;
  }
  const firstColor = readPaletteColor(bytes, paletteOffset);
  const secondColor = readPaletteColor(bytes, paletteOffset + 4);
  let blackIndex: number;
  if (isBlack(firstColor) && isWhite(secondColor)) {
    blackIndex = 0;
  } else if (isWhite(firstColor) && isBlack(secondColor)) {
    blackIndex = 1;
  } else {
    return undefined;
  }

  const blackPixels = new Array<boolean>(dib.width * dib.height);
  for (let y = 0; y < dib.height; y += 1) {
    const sourceRow = dib.bottomUp ? dib.height - 1 - y : y;
    const rowOffset = dib.bitsOffset + sourceRow * dib.stride;
    for (let x = 0; x < dib.width; x += 1) {
      const value = bytes[rowOffset + (x >> 3)];
      const colorIndex = (value >> (7 - (x & 7))) & 1;
      blackPixels[y * dib.width + x] = colorIndex === blackIndex;
    }
  }
  const brush: EmfPatternBrush = {
    handle: view.getUint32(offset + 8, true),
    creationOffset: offset,
    width: dib.width,
    height: dib.height,
    blackPixels,
  };
  return brush;
}

/** 创建同句柄的白色普通画刷，供已展开的图案蒙版安全占位。 */
export function createPatternBrushPlaceholder(handle: number) {
  const output = new Uint8Array(24);
  const view = createView(output);
  view.setUint32(0, EMR_CREATEBRUSHINDIRECT, true);
  view.setUint32(4, output.length, true);
  view.setUint32(8, handle, true);
  view.setUint32(12, 0, true);
  view.setUint32(16, 0x00ffffff, true);
  view.setUint32(20, 0, true);
  return output;
}

function positiveModulo(value: number, divisor: number) {
  return ((value % divisor) + divisor) % divisor;
}

/** 将黑白图案画刷的 DPa 位块运算展开为透明蒙版位图。 */
export function convertPatternAndBitBlt(
  bytes: Uint8Array,
  record: EmfRecord,
  context: PatternMaskContext,
) {
  if (record.size < 100) return undefined;
  const view = createView(bytes);
  const offset = record.offset;
  const destination = {
    x: view.getInt32(offset + 24, true),
    y: view.getInt32(offset + 28, true),
    width: view.getInt32(offset + 32, true),
    height: view.getInt32(offset + 36, true),
  };
  if (
    view.getUint32(offset + 40, true) !== ROP_DPA ||
    view.getUint32(offset + 84, true) !== 0 ||
    view.getUint32(offset + 88, true) !== 0 ||
    view.getUint32(offset + 92, true) !== 0 ||
    view.getUint32(offset + 96, true) !== 0 ||
    !isBitmapSizeSafe(destination.width, destination.height)
  ) {
    return undefined;
  }

  const pixels = new Uint8Array(destination.width * destination.height * 4);
  const { brush, originX, originY } = context;
  for (let y = 0; y < destination.height; y += 1) {
    const patternY = positiveModulo(destination.y + y - originY, brush.height);
    for (let x = 0; x < destination.width; x += 1) {
      const patternX = positiveModulo(destination.x + x - originX, brush.width);
      const output = (y * destination.width + x) * 4;
      const black = brush.blackPixels[patternY * brush.width + patternX];
      // D & 黑色等价于覆盖黑色；D & 白色等价于保持目标像素。
      pixels[output + 3] = black ? 255 : 1;
    }
  }
  return createStretchDibitsRecord(bytes, record, destination, {
    width: destination.width,
    height: destination.height,
    pixels,
  });
}
