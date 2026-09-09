/** Core 导出门面，不引入 Viewer UI。 */
export {
  OfficeExportService,
  createOfficeExportService,
  createOfficeExporterRegistry,
  exportOriginalOfficeFile,
} from '../export';
export type {
  OfficeExportError,
  OfficeExportErrorCode,
  OfficeExportRequest,
  OfficeExportResult,
  OfficeExporter,
  OfficeExporterRegistry,
  OfficeOriginalSource,
} from '../export';
