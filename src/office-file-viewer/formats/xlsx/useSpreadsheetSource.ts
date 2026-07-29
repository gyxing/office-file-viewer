import { useCallback, useEffect, useState, useSyncExternalStore } from 'react';
import type { SpreadsheetSource } from '../../services/spreadsheet/SpreadsheetSource';
import type { SpreadsheetSheet } from '../../services/spreadsheet/types';

/** 订阅 Source，并只为当前激活 Sheet 请求所需模型。 */
export function useSpreadsheetSource(
  source: SpreadsheetSource,
  activeSheetId: string | undefined,
) {
  const snapshot = useSyncExternalStore(
    source.subscribe.bind(source),
    source.getSnapshot.bind(source),
    source.getSnapshot.bind(source),
  );
  const activeDescriptor =
    snapshot.sheets.find((sheet) => sheet.id === activeSheetId) ??
    snapshot.sheets[0];
  const [retryRevision, setRetryRevision] = useState(0);
  const [activeSheet, setActiveSheet] = useState<SpreadsheetSheet>();
  const [loading, setLoading] = useState(Boolean(activeDescriptor));
  const [error, setError] = useState<Error>();
  const profile = activeDescriptor
    ? source.getProfile(activeDescriptor.id)
    : undefined;

  useEffect(() => {
    const sheetId = activeDescriptor?.id;
    if (!sheetId) {
      setActiveSheet(undefined);
      setLoading(false);
      setError(undefined);
      return undefined;
    }
    const controller = new AbortController();
    setActiveSheet(undefined);
    setLoading(true);
    setError(undefined);
    void source
      .ensureSheet(sheetId, controller.signal)
      .then(() =>
        source.getProfile(sheetId).gridMode === 'table'
          ? source.getMaterializedSheet(sheetId, controller.signal)
          : undefined,
      )
      .then(
        (sheet) => {
          if (controller.signal.aborted) return;
          setActiveSheet(sheet);
          setLoading(false);
        },
        (reason) => {
          if (controller.signal.aborted) return;
          setError(
            reason instanceof Error ? reason : new Error('工作表加载失败'),
          );
          setLoading(false);
        },
      );
    return () => controller.abort();
  }, [activeDescriptor?.id, retryRevision, source]);

  const retry = useCallback(() => {
    if (!activeDescriptor) return;
    source.retrySheet(activeDescriptor.id);
    setRetryRevision((value) => value + 1);
  }, [activeDescriptor, source]);

  return {
    snapshot,
    activeDescriptor,
    profile,
    activeSheet,
    loading,
    error,
    retry,
  };
}
