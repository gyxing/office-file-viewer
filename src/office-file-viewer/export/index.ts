/** 导出公共入口。 */
export {
  OfficeExportService,
  createOfficeExportService,
  createOfficeExporterRegistry,
  exportOriginalOfficeFile,
} from './OfficeExportService';
export type {
  OfficeExportError,
  OfficeExportErrorCode,
  OfficeExportRequest,
  OfficeExportResult,
  OfficeExporter,
  OfficeExporterRegistry,
  OfficeOriginalSource,
} from './types';
