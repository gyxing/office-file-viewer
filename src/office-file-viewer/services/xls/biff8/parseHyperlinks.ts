import type { OfficeHyperlink } from '../../../shared/hyperlink';
import type { SpreadsheetHyperlinkRange } from '../../spreadsheet/types';
import {
  createSpreadsheetExternalHyperlink,
  internalHyperlink,
} from '../../xlsx/parseXlsxHyperlinks';
import { Biff8Reader } from './Biff8Reader';

/** URL Moniker 的标准 COM 类标识，按 BIFF 小端字节序保存。 */
const URL_MONIKER_CLSID = 'E0C9EA79F9BACE118C8200AA004BA90B';
/** File Moniker 的标准 COM 类标识，按 BIFF 小端字节序保存。 */
const FILE_MONIKER_CLSID = '0303000000000000C000000000000046';

function bytesToHex(bytes: Uint8Array) {
  return Array.from(bytes, (value) => value.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

function readNullTerminatedUnicode(bytes: Uint8Array) {
  const view = new DataView(
    bytes.buffer,
    bytes.byteOffset,
    bytes.byteLength - (bytes.byteLength % 2),
  );
  let value = '';
  for (let offset = 0; offset + 1 < view.byteLength; offset += 2) {
    const code = view.getUint16(offset, true);
    if (!code) break;
    value += String.fromCharCode(code);
  }
  return value;
}

/** 读取 MS-OSHARED HyperlinkString，字符数包含结尾的空字符。 */
function readHyperlinkString(reader: Biff8Reader) {
  const characterCount = reader.readUint32();
  if (characterCount > Math.floor(reader.remaining / 2)) {
    throw new RangeError('BIFF8 HyperlinkString 长度超出记录边界');
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

function readFileMoniker(reader: Biff8Reader) {
  const parentCount = reader.readUint16();
  const pathLength = reader.readUint32();
  if (pathLength > reader.remaining) {
    throw new RangeError('BIFF8 FileMoniker 路径超出记录边界');
  }
  const path = readAnsiPath(reader.readBytes(pathLength));
  return `${'..\\'.repeat(parentCount)}${path}`;
}

function appendLocation(target: string, location: string | undefined) {
  if (!location) return target;
  return target.includes('#') ? target : `${target}#${location}`;
}

/** 解析 BIFF8 HLINK 记录中的单元格范围与 Hyperlink Object。 */
export function parseBiff8Hyperlink(
  data: Uint8Array,
): SpreadsheetHyperlinkRange | undefined {
  const reader = new Biff8Reader(data);
  const startRow = reader.readUint16() + 1;
  const endRow = reader.readUint16() + 1;
  const startColumn = reader.readUint16() + 1;
  const endColumn = reader.readUint16() + 1;
  reader.readBytes(16);
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
          throw new RangeError('BIFF8 URLMoniker 长度超出记录边界');
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

  const hyperlink: OfficeHyperlink | undefined = target
    ? createSpreadsheetExternalHyperlink(appendLocation(target, location))
    : location
    ? internalHyperlink(location)
    : undefined;
  if (!hyperlink) return undefined;
  return {
    ref: `${cellReference(startRow, startColumn)}:${cellReference(
      endRow,
      endColumn,
    )}`,
    startRow,
    endRow,
    startColumn,
    endColumn,
    hyperlink,
  };
}

/** 解析紧跟 HLINK 的 HLINKTOOLTIP 记录。 */
export function parseBiff8HyperlinkTooltip(data: Uint8Array) {
  const reader = new Biff8Reader(data);
  if (reader.readUint16() !== 0x0800) return undefined;
  const startRow = reader.readUint16() + 1;
  const endRow = reader.readUint16() + 1;
  const startColumn = reader.readUint16() + 1;
  const endColumn = reader.readUint16() + 1;
  const screenTip = readNullTerminatedUnicode(
    reader.readBytes(reader.remaining),
  );
  return screenTip
    ? {
        ref: `${cellReference(startRow, startColumn)}:${cellReference(
          endRow,
          endColumn,
        )}`,
        screenTip,
      }
    : undefined;
}

function cellReference(row: number, column: number) {
  let current = column;
  let label = '';
  while (current > 0) {
    current -= 1;
    label = String.fromCharCode(65 + (current % 26)) + label;
    current = Math.floor(current / 26);
  }
  return `${label}${row}`;
}
