import type { OfficeRelationship } from '../../shared/ooxml/media';
import type { OfficeArchiveReader } from '../../shared/ooxml/OfficeArchiveReader';
import type { OfficeTheme } from '../../shared/ooxml/theme';
import type { SpreadsheetPerformanceProfile } from '../spreadsheet/spreadsheetPerformance';
import type { SpreadsheetSheetDescriptor } from '../spreadsheet/SpreadsheetSource';
import type { StyleBook } from './parseXlsx';
import type { XlsxSharedStringSource } from './XlsxSharedStringSource';

/** XLSX Source 使用的 Sheet 描述符和包内路径。 */
export type XlsxSheetDescriptor = SpreadsheetSheetDescriptor & {
  /** 对应 worksheet/chartsheet XML 的未压缩字节数。 */
  sheetBytes: number;
  /** 当前 Sheet 的关系文件路径。 */
  relsPath: string;
  /** 当前 Sheet 的性能模式。 */
  performance: SpreadsheetPerformanceProfile;
};

/** XLSX 按需解析共享的包级结构。 */
export type XlsxPackageContext = {
  sessionId: string;
  reader: OfficeArchiveReader;
  relationships: Record<string, Record<string, OfficeRelationship>>;
  styles: StyleBook;
  theme: OfficeTheme;
  sharedStrings: XlsxSharedStringSource;
  descriptors: XlsxSheetDescriptor[];
};
