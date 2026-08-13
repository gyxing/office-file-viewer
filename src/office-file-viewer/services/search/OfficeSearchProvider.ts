import type { OfficeSearchProgressEmitter, OfficeSearchResult } from './types';

/** 搜索提供器单次占用主线程的目标时间片，单位为毫秒。 */
export const OFFICE_SEARCH_TIME_BUDGET_MS = 8;

function now() {
  return typeof performance === 'undefined' ? Date.now() : performance.now();
}

/** 创建与浏览器取消语义一致的搜索中止错误。 */
export function createOfficeSearchAbortError() {
  if (typeof DOMException !== 'undefined') {
    return new DOMException('Office 搜索已取消', 'AbortError');
  }
  const error = new Error('Office 搜索已取消');
  error.name = 'AbortError';
  return error;
}

/** 在可取消的扫描步骤前统一检查查询状态。 */
export function throwIfOfficeSearchAborted(signal: AbortSignal) {
  if (signal.aborted) throw createOfficeSearchAbortError();
}

/** 判断异常是否由主动取消当前搜索产生。 */
export function isOfficeSearchAbortError(error: unknown) {
  return error instanceof Error && error.name === 'AbortError';
}

function yieldToBrowser(signal: AbortSignal) {
  throwIfOfficeSearchAborted(signal);
  return new Promise<void>((resolve, reject) => {
    let timer: ReturnType<typeof setTimeout>;
    const abort = () => {
      clearTimeout(timer);
      signal.removeEventListener('abort', abort);
      reject(createOfficeSearchAbortError());
    };
    timer = setTimeout(() => {
      signal.removeEventListener('abort', abort);
      resolve();
    }, 0);
    signal.addEventListener('abort', abort, { once: true });
  });
}

/**
 * 汇集提供器产生的结果，并在时间片用尽后输出增量批次和让出主线程。
 */
export class OfficeSearchBatchWriter {
  private readonly pendingItems: OfficeSearchResult[] = [];
  private scanned = 0;
  private total: number;
  private sliceStartedAt = now();

  constructor(
    private readonly emit: OfficeSearchProgressEmitter,
    private readonly signal: AbortSignal,
    total: number,
  ) {
    this.total = total;
  }

  /** 更新渐进数据源当前已知的扫描总量。 */
  setTotal(total: number) {
    this.total = Math.max(this.total, total);
  }

  /** 记录单个扫描单元及其零个或多个匹配结果。 */
  async append(items: readonly OfficeSearchResult[], scannedUnits = 1) {
    throwIfOfficeSearchAborted(this.signal);
    this.pendingItems.push(...items);
    this.scanned += Math.max(0, scannedUnits);
    if (now() - this.sliceStartedAt < OFFICE_SEARCH_TIME_BUDGET_MS) return;
    this.flush(false);
    await yieldToBrowser(this.signal);
    this.sliceStartedAt = now();
  }

  /** 完成扫描并保证至少输出一个 complete 批次。 */
  complete() {
    throwIfOfficeSearchAborted(this.signal);
    this.flush(true);
  }

  private flush(complete: boolean) {
    if (!complete && !this.pendingItems.length) return;
    this.emit({
      items: this.pendingItems.splice(0),
      scanned: this.scanned,
      total: Math.max(this.total, this.scanned),
      complete,
    });
  }
}
