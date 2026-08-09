/** OfficeArt 形状超链接解析后的宿主无关目标。 */
export type OfficeArtHyperlinkTarget = {
  /** 外部地址或文件路径。 */
  target?: string;
  /** 同一文档内的位置或附加片段。 */
  location?: string;
};

/** 标准 IHlink 的类标识，按复合文档小端字节顺序保存。 */
const STANDARD_HLINK_CLSID = 'D0C9EA79F9BACE118C8200AA004BA90B';
/** URL Moniker 的标准类标识。 */
const URL_MONIKER_CLSID = 'E0C9EA79F9BACE118C8200AA004BA90B';
/** File Moniker 的标准类标识。 */
const FILE_MONIKER_CLSID = '0303000000000000C000000000000046';

/** 在 pihlShape 复合属性边界内顺序读取 IHlink 字段。 */
class OfficeArtHyperlinkReader {
  private offset = 0;

  constructor(private readonly data: Uint8Array) {}

  get remaining() {
    return this.data.length - this.offset;
  }

  readUint16() {
    this.ensure(2);
    const value = new DataView(
      this.data.buffer,
      this.data.byteOffset + this.offset,
      2,
    ).getUint16(0, true);
    this.offset += 2;
    return value;
  }

  readUint32() {
    this.ensure(4);
    const value = new DataView(
      this.data.buffer,
      this.data.byteOffset + this.offset,
      4,
    ).getUint32(0, true);
    this.offset += 4;
    return value;
  }

  readBytes(length: number) {
    this.ensure(length);
    const value = this.data.slice(this.offset, this.offset + length);
    this.offset += length;
    return value;
  }

  private ensure(length: number) {
    if (length < 0 || length > this.remaining) {
      throw new RangeError('OfficeArt 超链接数据超出属性边界');
    }
  }
}

function bytesToHex(bytes: Uint8Array) {
  return Array.from(bytes, (value) => value.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

function readNullTerminatedUnicode(bytes: Uint8Array) {
  const evenLength = bytes.byteLength - (bytes.byteLength % 2);
  return new TextDecoder('utf-16le')
    .decode(bytes.slice(0, evenLength))
    .split('\u0000')[0];
}

function readHyperlinkString(reader: OfficeArtHyperlinkReader) {
  const characterCount = reader.readUint32();
  if (characterCount > Math.floor(reader.remaining / 2)) {
    throw new RangeError('OfficeArt 超链接字符串长度越界');
  }
  return readNullTerminatedUnicode(reader.readBytes(characterCount * 2));
}

function readAnsiPath(bytes: Uint8Array) {
  let value = '';
  for (const byte of bytes) {
    if (!byte) break;
    value += String.fromCharCode(byte);
  }
  return value;
}

function readFileMoniker(reader: OfficeArtHyperlinkReader) {
  const parentCount = reader.readUint16();
  const pathLength = reader.readUint32();
  if (pathLength > reader.remaining) {
    throw new RangeError('OfficeArt FileMoniker 路径长度越界');
  }
  return `${'..\\'.repeat(parentCount)}${readAnsiPath(
    reader.readBytes(pathLength),
  )}`;
}

function readHyperlinkObject(reader: OfficeArtHyperlinkReader) {
  if (reader.readUint32() !== 2) return undefined;
  const flags = reader.readUint32();
  if (flags & 0x10) readHyperlinkString(reader);
  if (flags & 0x80) readHyperlinkString(reader);

  let target: string | undefined;
  if (flags & 0x01) {
    if (flags & 0x100) {
      target = readHyperlinkString(reader);
    } else {
      const monikerClsid = bytesToHex(reader.readBytes(16));
      if (monikerClsid === URL_MONIKER_CLSID) {
        const byteLength = reader.readUint32();
        if (byteLength > reader.remaining) {
          throw new RangeError('OfficeArt URLMoniker 长度越界');
        }
        target = readNullTerminatedUnicode(reader.readBytes(byteLength));
      } else if (monikerClsid === FILE_MONIKER_CLSID) {
        target = readFileMoniker(reader);
      } else {
        return undefined;
      }
    }
  }
  const location = flags & 0x08 ? readHyperlinkString(reader) : undefined;
  if (flags & 0x20) reader.readBytes(16);
  if (flags & 0x40) reader.readBytes(8);
  return target || location ? { target, location } : undefined;
}

/** 解析 pihlShape_complex 中的标准 IHlink，不执行宏或未知 Moniker。 */
export function parseOfficeArtHyperlink(
  data: Uint8Array,
): OfficeArtHyperlinkTarget | undefined {
  const reader = new OfficeArtHyperlinkReader(data);
  if (bytesToHex(reader.readBytes(16)) !== STANDARD_HLINK_CLSID) {
    return undefined;
  }
  return readHyperlinkObject(reader);
}
