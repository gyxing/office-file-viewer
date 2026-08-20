import type { OfficeSearchProvider } from '../../../search/types';
import type { SpreadsheetAnnotation } from '../../../spreadsheet/semantics/types';
import type { SpreadsheetPerformanceProfile } from '../../../spreadsheet/spreadsheetPerformance';
import type {
  SpreadsheetSheetLayout,
  SpreadsheetSource,
  SpreadsheetSourceSnapshot,
} from '../../../spreadsheet/SpreadsheetSource';
import type {
  SpreadsheetRange,
  SpreadsheetRangeData,
  SpreadsheetSheet,
} from '../../../spreadsheet/types';
import type { WorkerSpreadsheetSourceState } from '../../protocol/messages';
import { WorkerSourceClient } from './WorkerSourceClient';

/** 将 Worker 中长期持有的 XLSX Source 适配为现有虚拟表格接口。 */
export class WorkerSpreadsheetSource implements SpreadsheetSource {
  readonly searchProvider: OfficeSearchProvider = {
    kind: 'spreadsheet',
    search: (query, emit, signal) => this.client.search(query, emit, signal),
  };
  private readonly listeners = new Set<() => void>();
  private readonly unsubscribeUpdate: () => void;
  private readonly unsubscribeFailure: () => void;
  private snapshot: SpreadsheetSourceSnapshot;
  private profiles: Record<string, SpreadsheetPerformanceProfile>;
  private layouts: Record<string, SpreadsheetSheetLayout>;
  private disposed = false;

  constructor(
    private readonly client: WorkerSourceClient,
    initial: WorkerSpreadsheetSourceState,
  ) {
    this.snapshot = initial.snapshot;
    this.profiles = initial.profiles;
    this.layouts = initial.layouts;
    this.unsubscribeUpdate = client.subscribe((source) => {
      if (source.kind !== 'xlsx') return;
      this.snapshot = source.snapshot;
      this.profiles = source.profiles;
      this.layouts = source.layouts;
      this.emitChange();
    });
    this.unsubscribeFailure = client.subscribeFailure((error) => {
      this.snapshot = {
        ...this.snapshot,
        revision: this.snapshot.revision + 1,
        sheets: this.snapshot.sheets.map((sheet) =>
          sheet.status === 'estimated'
            ? {
                ...sheet,
                revision: sheet.revision + 1,
                status: 'error',
                errorMessage: error.message,
              }
            : sheet,
        ),
      };
      this.emitChange();
    });
  }

  getSnapshot = () => this.snapshot;

  subscribe = (listener: () => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  getProfile(sheetId: string) {
    const profile = this.profiles[sheetId];
    if (!profile) throw new RangeError(`工作表不存在：${sheetId}`);
    return profile;
  }

  getSheetLayout(sheetId: string) {
    const layout = this.layouts[sheetId];
    if (!layout) throw new RangeError(`工作表不存在：${sheetId}`);
    return layout;
  }

  ensureSheet(sheetId: string, signal?: AbortSignal) {
    return this.client.request<void>('ensure-sheet', { sheetId }, { signal });
  }

  getMaterializedSheet(sheetId: string, signal?: AbortSignal) {
    return this.client.request<SpreadsheetSheet | undefined>(
      'get-materialized-sheet',
      { sheetId },
      { signal },
    );
  }

  getAnnotations(sheetId: string, signal?: AbortSignal) {
    return this.client.request<readonly SpreadsheetAnnotation[]>(
      'get-sheet-annotations',
      { sheetId },
      { signal },
    );
  }

  getRange(sheetId: string, range: SpreadsheetRange, signal?: AbortSignal) {
    return this.client.request<SpreadsheetRangeData>(
      'get-range',
      { sheetId, range },
      { signal },
    );
  }

  retainRange(sheetId: string, range: SpreadsheetRange) {
    return this.client.retain('retain-sheet-range', { sheetId, range });
  }

  retrySheet(sheetId: string) {
    void this.client.request('retry-sheet', { sheetId }).catch(() => undefined);
  }

  async dispose() {
    if (this.disposed) return;
    this.disposed = true;
    this.unsubscribeUpdate();
    this.unsubscribeFailure();
    this.listeners.clear();
    await this.client.dispose();
  }

  private emitChange() {
    this.listeners.forEach((listener) => listener());
  }
}
