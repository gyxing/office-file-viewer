// OfficeArt 公共入口，集中导出记录解析器、错误类型和标准化图形资源模型。
export { OFFICE_ART_RECORD } from './constants';
export { OfficeArtParseError } from './OfficeArtParseError';
export { parseOfficeArtRecords } from './parseOfficeArt';
export type {
  DecodedBitmap,
  OfficeArtImageFormat,
  OfficeArtRecord,
  OfficeArtWarning,
  ParsedOfficeArtBlip,
} from './types';
