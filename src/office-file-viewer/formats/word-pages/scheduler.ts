/** 在长任务边界统一响应取消信号。 */
export function throwIfAborted(signal?: AbortSignal) {
  if (!signal?.aborted) return;
  const error = new Error('操作已取消');
  error.name = 'AbortError';
  throw error;
}

/** 优先使用原生 scheduler.yield，并提供浏览器和 SSR 降级。 */
export async function yieldToMainThread(signal?: AbortSignal) {
  throwIfAborted(signal);
  const browserScheduler = (
    globalThis as typeof globalThis & {
      scheduler?: { yield?(): Promise<void> };
    }
  ).scheduler;
  if (browserScheduler?.yield) {
    await browserScheduler.yield();
  } else if (typeof requestAnimationFrame === 'function') {
    await new Promise<void>((resolve) => {
      requestAnimationFrame(() => resolve());
    });
  } else {
    await new Promise<void>((resolve) => {
      setTimeout(resolve, 0);
    });
  }
  throwIfAborted(signal);
}

/** 创建可重复复位的主线程时间预算。 */
export function createTimeBudget(milliseconds = 8) {
  let startedAt = performance.now();
  return {
    shouldYield: () => performance.now() - startedAt >= milliseconds,
    reset: () => {
      startedAt = performance.now();
    },
  };
}
