import type { RefObject } from 'react';
import { useEffect, useState } from 'react';
import type { OfficeFileViewerZoomMode } from './viewState';

/** 内容超出视口的方向。 */
export type OfficeOverflowAxis = 'horizontal' | 'vertical' | 'both';

/** 视口溢出检测所需的预览状态。 */
type UseOfficeOverflowNoticeOptions = {
  /** OfficeFileViewer 根容器。 */
  containerRef: RefObject<HTMLElement>;
  /** 文件会话变化时重新显示提示。 */
  sessionKey?: string;
  /** 工作表或幻灯片切换时重新显示提示。 */
  activeKey?: string | number;
  /** 当前缩放百分比。 */
  zoom: number;
  /** 自适应模式下由外壳自动消除可见溢出，不显示提示。 */
  zoomMode: OfficeFileViewerZoomMode;
  /** 没有可渲染内容时关闭检测。 */
  enabled?: boolean;
};

/** 抵消滚动条取整造成的 1～2 像素误差，避免把贴边内容误报为溢出。 */
const OVERFLOW_EPSILON = 2;
/** 仅观察格式查看器明确声明的真实滚动视口。 */
const VIEWPORT_SELECTOR =
  '[data-office-fit-viewport="true"][data-office-overflow-viewport="true"]';

function readOverflowAxis(
  viewport: HTMLElement,
): OfficeOverflowAxis | undefined {
  if (!viewport.getClientRects().length) return undefined;
  const horizontal =
    viewport.scrollWidth > viewport.clientWidth + OVERFLOW_EPSILON;
  const verticalEnabled = viewport.dataset.officeOverflowVertical !== 'false';
  const vertical =
    verticalEnabled &&
    viewport.scrollHeight > viewport.clientHeight + OVERFLOW_EPSILON;
  if (horizontal && vertical) return 'both';
  if (horizontal) return 'horizontal';
  if (vertical) return 'vertical';
  return undefined;
}

function mergeOverflowAxis(
  current: OfficeOverflowAxis | undefined,
  next: OfficeOverflowAxis | undefined,
) {
  if (!current) return next;
  if (!next || current === next) return current;
  return 'both' as const;
}

/** 观察真实滚动尺寸，向用户提示需要横向或纵向滚动的内容。 */
export function useOfficeOverflowNotice({
  containerRef,
  sessionKey,
  activeKey,
  zoom,
  zoomMode,
  enabled = true,
}: UseOfficeOverflowNoticeOptions) {
  const [axis, setAxis] = useState<OfficeOverflowAxis>();
  const [dismissed, setDismissed] = useState(false);

  useEffect(() => {
    const container = containerRef.current;
    const shouldObserve = enabled && zoomMode === 'percentage';
    if (!container || !shouldObserve) {
      setAxis(undefined);
      setDismissed(false);
      return undefined;
    }

    let disposed = false;
    let frame: number | undefined;
    let timeout: ReturnType<typeof setTimeout> | undefined;
    const deferredMeasurements: ReturnType<typeof setTimeout>[] = [];
    const candidates = new Set<HTMLElement>();

    const measure = () => {
      frame = undefined;
      timeout = undefined;
      if (disposed) return;
      let nextAxis: OfficeOverflowAxis | undefined;
      candidates.forEach((candidate) => {
        nextAxis = mergeOverflowAxis(nextAxis, readOverflowAxis(candidate));
      });
      setAxis((current) => (current === nextAxis ? current : nextAxis));
    };
    const scheduleMeasure = () => {
      if (disposed || frame !== undefined || timeout !== undefined) return;
      if (typeof window !== 'undefined' && window.requestAnimationFrame) {
        frame = window.requestAnimationFrame(measure);
      } else {
        timeout = setTimeout(measure, 0);
      }
    };
    const dismiss = () => {
      if (!disposed) setDismissed(true);
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (
        event.key === 'ArrowDown' ||
        event.key === 'ArrowLeft' ||
        event.key === 'ArrowRight' ||
        event.key === 'ArrowUp' ||
        event.key === 'End' ||
        event.key === 'Home' ||
        event.key === 'PageDown' ||
        event.key === 'PageUp' ||
        event.key === ' '
      ) {
        dismiss();
      }
    };
    const attach = (candidate: HTMLElement) => {
      if (candidates.has(candidate)) return;
      candidates.add(candidate);
      // 只监听用户输入意图，不监听 scroll 事件，避免脚本 scrollTo 复位时误关闭提示。
      candidate.addEventListener('wheel', dismiss, { passive: true });
      candidate.addEventListener('touchstart', dismiss, { passive: true });
      candidate.addEventListener('pointerdown', dismiss, { passive: true });
      candidate.addEventListener('keydown', handleKeyDown);
      scheduleMeasure();
    };
    const detach = (candidate: HTMLElement) => {
      if (!candidates.delete(candidate)) return;
      candidate.removeEventListener('wheel', dismiss);
      candidate.removeEventListener('touchstart', dismiss);
      candidate.removeEventListener('pointerdown', dismiss);
      candidate.removeEventListener('keydown', handleKeyDown);
    };
    const refreshCandidates = () => {
      const nextCandidates = new Set(
        Array.from(container.querySelectorAll<HTMLElement>(VIEWPORT_SELECTOR)),
      );
      candidates.forEach((candidate) => {
        if (!nextCandidates.has(candidate)) detach(candidate);
      });
      nextCandidates.forEach(attach);
      scheduleMeasure();
    };

    setAxis(undefined);
    setDismissed(false);
    refreshCandidates();
    // 翻页可能只替换画布内容而不改变滚动容器尺寸，延后两次测量覆盖异步渲染提交。
    scheduleMeasure();
    deferredMeasurements.push(
      setTimeout(scheduleMeasure, 80),
      setTimeout(scheduleMeasure, 240),
    );

    const resizeObserver =
      typeof ResizeObserver === 'undefined'
        ? undefined
        : new ResizeObserver(scheduleMeasure);
    resizeObserver?.observe(container);
    candidates.forEach((candidate) => resizeObserver?.observe(candidate));
    const mutationObserver =
      typeof MutationObserver === 'undefined'
        ? undefined
        : new MutationObserver(() => {
            refreshCandidates();
            candidates.forEach((candidate) =>
              resizeObserver?.observe(candidate),
            );
          });
    mutationObserver?.observe(container, { childList: true, subtree: true });
    window.addEventListener('resize', scheduleMeasure);

    return () => {
      disposed = true;
      if (frame !== undefined) window.cancelAnimationFrame(frame);
      if (timeout !== undefined) clearTimeout(timeout);
      deferredMeasurements.forEach((measurement) => clearTimeout(measurement));
      Array.from(candidates).forEach(detach);
      resizeObserver?.disconnect();
      mutationObserver?.disconnect();
      window.removeEventListener('resize', scheduleMeasure);
    };
  }, [activeKey, containerRef, enabled, sessionKey, zoom, zoomMode]);

  return axis && !dismissed ? axis : undefined;
}
