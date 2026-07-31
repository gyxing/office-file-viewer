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
  /** 当前解析或预览会话的标识。 */
  sessionId: string;
  /** 用于按需读取源数据的读取器。 */
  reader: OfficeArchiveReader;
  /** 按关系文件路径组织的 OOXML 关系映射。 */
  relationships: Record<string, Record<string, OfficeRelationship>>;
  /** 工作簿共享的字体、填充、边框和单元格格式。 */
  styles: StyleBook;
  /** 当前文档使用的主题颜色和字体配置。 */
  theme: OfficeTheme;
  /** 支持按索引批量读取的共享字符串源。 */
  sharedStrings: XlsxSharedStringSource;
  /** 按源顺序排列的轻量描述信息。 */
  descriptors: XlsxSheetDescriptor[];
};
