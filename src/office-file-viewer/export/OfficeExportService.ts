import { OfficeFileViewerError } from '../services/errors/OfficeFileViewerError';
import { OFFICE_FORMAT_METADATA } from '../services/parsing/formatDefinitions';
import type {
  OfficeExportRequest,
  OfficeExportResult,
  OfficeExporter,
  OfficeExporterRegistry,
  OfficeOriginalSource,
} from './types';

function normalizeFormat(format: unknown) {
  if (typeof format !== 'string') return '';
  const value = format.trim().toLowerCase();
  return value.startsWith('.') ? value.slice(1) : value;
}

function isFileSource(source: OfficeOriginalSource): source is File {
  return typeof File !== 'undefined' && source instanceof File;
}

function isBlobSource(source: OfficeOriginalSource): source is Blob {
  return typeof Blob !== 'undefined' && source instanceof Blob;
}

function createExportError(
  code:
    | 'EXPORT_UNSUPPORTED'
    | 'EXPORT_CANCELLED'
    | 'EXPORT_FILENAME_REQUIRED'
    | 'EXPORT_RESOURCE_FAILED'
    | 'EXPORT_WRITE_FAILED',
  message: string,
  format?: string,
  cause?: unknown,
) {
  return new OfficeFileViewerError(code, message, {
    stage: 'export',
    format,
    recoverable: code === 'EXPORT_CANCELLED',
    cause,
  });
}

function ensureNotAborted(signal?: AbortSignal, format?: string) {
  if (signal?.aborted) {
    throw createExportError('EXPORT_CANCELLED', '导出已取消', format);
  }
}

function validateExportResult(
  result: OfficeExportResult,
  format: string,
): OfficeExportResult {
  if (
    !result ||
    typeof result !== 'object' ||
    !result.blob ||
    typeof result.blob.arrayBuffer !== 'function' ||
    typeof result.fileName !== 'string' ||
    !result.fileName.trim() ||
    typeof result.format !== 'string' ||
    !result.format.trim()
  ) {
    throw createExportError(
      'EXPORT_WRITE_FAILED',
      '导出器返回了无效结果',
      format,
    );
  }
  return result;
}

function getFormatExtension(format: string) {
  const normalized = normalizeFormat(format);
  for (const metadata of Object.values(OFFICE_FORMAT_METADATA)) {
    const extension = metadata.extensions.find(
      (candidate) => normalizeFormat(candidate) === normalized,
    );
    if (extension) return extension.slice(1);
  }
  return normalized;
}

function inferFileName(source: OfficeOriginalSource, format?: string) {
  if (isFileSource(source) && source.name) return source.name;
  if (!format) return undefined;
  return `office-export.${getFormatExtension(format)}`;
}

function inferMimeType(source: OfficeOriginalSource, format?: string) {
  if (isFileSource(source) && source.type) return source.type;
  if (!format) return undefined;
  const normalized = normalizeFormat(format);
  for (const metadata of Object.values(OFFICE_FORMAT_METADATA)) {
    if (
      metadata.extensions.some(
        (extension) => normalizeFormat(extension) === normalized,
      )
    ) {
      return metadata.mimeTypes.find(
        ({ extension }) => normalizeFormat(extension) === normalized,
      )?.mimeType;
    }
  }
  return undefined;
}

async function toBlob(
  source: OfficeOriginalSource,
  type: string | undefined,
  signal?: AbortSignal,
) {
  ensureNotAborted(signal);
  if (isBlobSource(source))
    return source.slice(0, source.size, type ?? source.type);
  if (source instanceof ArrayBuffer) return new Blob([source], { type });
  const buffer = source.buffer.slice(
    source.byteOffset,
    source.byteOffset + source.byteLength,
  );
  return new Blob([buffer], { type });
}

/** 创建导出器注册表，重复格式会被拒绝以保持结果确定。 */
export function createOfficeExporterRegistry(
  exporters: readonly OfficeExporter[] = [],
): OfficeExporterRegistry {
  if (!Array.isArray(exporters)) {
    throw createExportError('EXPORT_WRITE_FAILED', '导出器列表无效');
  }
  const byFormat = new Map<string, OfficeExporter>();
  let disposed = false;

  const register = (exporter: OfficeExporter) => {
    if (
      disposed ||
      !exporter ||
      typeof exporter.id !== 'string' ||
      !exporter.id.trim() ||
      !Array.isArray(exporter.formats) ||
      typeof exporter.export !== 'function'
    ) {
      throw createExportError('EXPORT_WRITE_FAILED', '导出器定义无效');
    }
    if (
      !exporter.formats.length ||
      exporter.formats.some((format) => typeof format !== 'string')
    ) {
      throw createExportError('EXPORT_WRITE_FAILED', '导出器未声明格式');
    }
    const formats = exporter.formats.map(normalizeFormat);
    if (!formats.length || formats.some((format) => !format)) {
      throw createExportError('EXPORT_WRITE_FAILED', '导出器格式为空');
    }
    if (new Set(formats).size !== formats.length) {
      throw createExportError('EXPORT_WRITE_FAILED', '导出器声明了重复格式');
    }
    if (formats.some((format) => byFormat.has(format))) {
      throw createExportError('EXPORT_WRITE_FAILED', '导出格式已注册');
    }
    formats.forEach((format) => byFormat.set(format, exporter));
    let active = true;
    return () => {
      if (!active) return;
      active = false;
      formats.forEach((format) => {
        if (byFormat.get(format) === exporter) byFormat.delete(format);
      });
    };
  };

  const unregisters: Array<() => void> = [];
  try {
    exporters.forEach((exporter) => {
      unregisters.push(register(exporter));
    });
  } catch (error) {
    unregisters
      .splice(0)
      .reverse()
      .forEach((unregister) => unregister());
    byFormat.clear();
    throw error;
  }
  return {
    list: () => Object.freeze([...new Set(byFormat.values())]),
    getByFormat: (format) => byFormat.get(normalizeFormat(format)),
    register,
    dispose() {
      if (disposed) return;
      disposed = true;
      unregisters.splice(0).forEach((unregister) => unregister());
      byFormat.clear();
    },
  };
}

/**
 * 导出服务，默认只支持原始文件复制；修改后写回由 Exporter 提供。
 * @experimental 导出服务首版。
 */
export class OfficeExportService {
  private readonly registry: OfficeExporterRegistry;
  private readonly ownsRegistry: boolean;

  constructor(registry?: OfficeExporterRegistry) {
    this.registry = registry ?? createOfficeExporterRegistry();
    this.ownsRegistry = !registry;
  }

  async exportDocument(
    request: OfficeExportRequest,
  ): Promise<OfficeExportResult> {
    if (!request || !request.document) {
      throw createExportError('EXPORT_WRITE_FAILED', '导出请求缺少文档快照');
    }
    const format = normalizeFormat(request.format ?? request.document.format);
    if (!format) {
      throw createExportError('EXPORT_UNSUPPORTED', '导出格式为空');
    }
    ensureNotAborted(request.signal, format);
    const exporter = this.registry.getByFormat(format);
    if (!exporter) {
      throw createExportError(
        'EXPORT_UNSUPPORTED',
        `暂不支持导出格式：${format}`,
        format,
      );
    }
    try {
      const result = await exporter.export({ ...request, format });
      ensureNotAborted(request.signal, format);
      return validateExportResult(result, format);
    } catch (error) {
      if (
        request.signal?.aborted ||
        (error instanceof Error && error.name === 'AbortError')
      ) {
        throw createExportError(
          'EXPORT_CANCELLED',
          '导出已取消',
          format,
          error,
        );
      }
      if (error instanceof OfficeFileViewerError) {
        if (error.stage === 'resource') {
          throw createExportError(
            'EXPORT_RESOURCE_FAILED',
            error.message,
            format,
            error,
          );
        }
        throw error;
      }
      throw createExportError(
        'EXPORT_WRITE_FAILED',
        '导出文件失败',
        format,
        error,
      );
    }
  }

  async exportOriginalOfficeFile(
    source: OfficeOriginalSource,
    options: { fileName?: string; format?: string; signal?: AbortSignal } = {},
  ): Promise<OfficeExportResult> {
    const exportOptions = options ?? {};
    const format = exportOptions.format
      ? normalizeFormat(exportOptions.format)
      : isFileSource(source)
      ? normalizeFormat(source.name.split('.').pop() ?? '')
      : exportOptions.fileName
      ? normalizeFormat(exportOptions.fileName.split('.').pop() ?? '')
      : undefined;
    ensureNotAborted(exportOptions.signal, format);
    const fileName = exportOptions.fileName ?? inferFileName(source, format);
    if (!fileName) {
      throw createExportError(
        'EXPORT_FILENAME_REQUIRED',
        '原始文件导出需要文件名或格式',
        format,
      );
    }
    try {
      const blob = await toBlob(
        source,
        inferMimeType(source, format),
        exportOptions.signal,
      );
      return {
        blob,
        fileName,
        format: format ?? normalizeFormat(fileName.split('.').pop() ?? ''),
      };
    } catch (error) {
      if (error instanceof OfficeFileViewerError) throw error;
      throw createExportError(
        'EXPORT_WRITE_FAILED',
        '原始文件导出失败',
        format,
        error,
      );
    }
  }

  dispose() {
    if (this.ownsRegistry) this.registry.dispose();
  }
}

/** 创建独立的导出服务实例。 */
export function createOfficeExportService(options?: {
  exporters?: readonly OfficeExporter[];
}) {
  const registry = createOfficeExporterRegistry(options?.exporters);
  const service = new OfficeExportService(registry);
  return {
    exportDocument: (request: OfficeExportRequest) =>
      service.exportDocument(request),
    exportOriginalOfficeFile: (
      source: OfficeOriginalSource,
      exportOptions?: {
        fileName?: string;
        format?: string;
        signal?: AbortSignal;
      },
    ) => service.exportOriginalOfficeFile(source, exportOptions),
    dispose: () => {
      registry.dispose();
    },
  };
}

/** 便捷导出原始文件，不要求宿主持有服务实例。 */
export function exportOriginalOfficeFile(
  source: OfficeOriginalSource,
  options?: { fileName?: string; format?: string; signal?: AbortSignal },
): Promise<OfficeExportResult> {
  const service = new OfficeExportService();
  return service
    .exportOriginalOfficeFile(source, options)
    .finally(() => service.dispose());
}
