import type {
  OfficeFileViewerImagePreviewConfig,
  ResolvedOfficeImagePreviewOptions,
} from './types';

/** 将公共图片预览配置转换为没有可选字段的内部配置。 */
export function resolveOfficeImagePreviewOptions(
  config: OfficeFileViewerImagePreviewConfig | undefined,
): ResolvedOfficeImagePreviewOptions {
  if (config === false) {
    return { enabled: false, download: false, contextMenu: false };
  }
  if (config === true || config === undefined) {
    return { enabled: true, download: true, contextMenu: true };
  }
  return {
    enabled: true,
    download: config.download !== false,
    contextMenu: config.contextMenu !== false,
  };
}
