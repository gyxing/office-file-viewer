import { getOfficeCapabilities } from '../../core/model';
import type {
  OfficeCapabilities,
  OfficeDocumentRuntime,
} from '../../core/types';
import type { OfficeViewerMeta } from './useOfficeViewerController';

/** 根据 Viewer 当前会话和控制器元数据生成能力快照。 */
export function getOfficeViewerCapabilities(
  meta: OfficeViewerMeta,
  customRuntime?: OfficeDocumentRuntime,
): OfficeCapabilities | undefined {
  if (!meta.hasRenderableContent) return undefined;
  if (customRuntime) return customRuntime.snapshot.capabilities;
  if (!meta.previewKind) return undefined;
  return {
    ...getOfficeCapabilities(meta.previewKind),
    canFullscreen: meta.fullscreenSupported,
    canExportOriginal: Boolean(meta.currentFile),
  };
}
