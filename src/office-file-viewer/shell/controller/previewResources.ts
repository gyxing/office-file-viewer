import type { OfficeFileViewerPreviewState } from '../../services/parsing/internalTypes';
import { disposeParsedOfficeFile } from '../../services/preview';
import { disposeDocumentSession } from '../../services/session';

/**
 * 释放查看器当前持有的预览资源。
 * 同一会话的渐进快照共享资源，因此 partial 替换时不释放；仅在会话被替换、失败放弃或组件卸载时释放最新所有者。
 */
export async function disposeViewerPreviewState(
  preview: OfficeFileViewerPreviewState | undefined,
): Promise<void> {
  if (!preview) return;
  if (preview.mode === 'materialized') {
    await disposeParsedOfficeFile(preview.model);
    return;
  }
  await disposeDocumentSession(preview.source);
}
