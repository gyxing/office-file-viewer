import type { RefObject } from 'react';
import { useEffect } from 'react';
import type { PreviewKind } from '../services/preview';
import { isSpreadsheetPreviewKind } from '../services/preview';
import { normalizeOfficeZoom } from './normalizeOfficeZoom';
import type { OfficeFileViewerZoomMode } from './viewState';

/** 自适应缩放测量所需的稳定视图参数。 */
type UseOfficeFitZoomOptions = {
  /** 当前预览器根元素。 */
  containerRef: RefObject<HTMLDivElement>;
  /** 当前固定比例或自适应缩放模式。 */
  mode: OfficeFileViewerZoomMode;
  /** 当前已经应用的缩放比例。 */
  zoom: number;
  /** 当前文件的内部预览格式。 */
  previewKind?: PreviewKind;
  /** 文件切换时用于重建观察器的会话标识。 */
  sessionKey?: string;
  /** 幻灯片或工作表切换时用于重新测量的活动标识。 */
  activeKey?: string | number;
  /** 把测量结果写回控制器，同时保持自适应模式。 */
  onZoom(zoom: number): void;
};

/** 读取由格式查看器写入 data 属性的正数尺寸。 */
function readPositiveDimension(element: HTMLElement, name: string) {
  const value = Number(element.dataset[name]);
  return Number.isFinite(value) && value > 0 ? value : undefined;
}

/** 读取视口内边距，避免页面与滚动容器边缘贴合。 */
function getViewportContentSize(viewport: HTMLElement) {
  const style = window.getComputedStyle(viewport);
  const horizontalPadding =
    Number.parseFloat(style.paddingLeft) +
    Number.parseFloat(style.paddingRight);
  const verticalPadding =
    Number.parseFloat(style.paddingTop) +
    Number.parseFloat(style.paddingBottom);
  return {
    width: Math.max(0, viewport.clientWidth - horizontalPadding),
    height: Math.max(0, viewport.clientHeight - verticalPadding),
  };
}

/** 根据格式查看器声明的基础尺寸计算适应宽度或适应页面比例。 */
function measureFitZoom(
  container: HTMLElement,
  mode: Exclude<OfficeFileViewerZoomMode, 'percentage'>,
  previewKind: PreviewKind,
) {
  const viewport = container.querySelector<HTMLElement>(
    '[data-office-fit-viewport="true"]',
  );
  const target = container.querySelector<HTMLElement>(
    '[data-office-fit-target="true"]',
  );
  if (!viewport || !target) return undefined;

  const baseWidth = readPositiveDimension(target, 'officeFitBaseWidth');
  const baseHeight = readPositiveDimension(target, 'officeFitBaseHeight');
  if (!baseWidth || !baseHeight) return undefined;
  const fixedWidth = readPositiveDimension(target, 'officeFitFixedWidth') ?? 0;
  const fixedHeight =
    readPositiveDimension(target, 'officeFitFixedHeight') ?? 0;
  const available = getViewportContentSize(viewport);
  const widthZoom = ((available.width - fixedWidth) / baseWidth) * 100;
  const spreadsheet = isSpreadsheetPreviewKind(previewKind);
  const rawZoom =
    mode === 'fit-page' && !spreadsheet
      ? Math.min(
          widthZoom,
          ((available.height - fixedHeight) / baseHeight) * 100,
        )
      : widthZoom;
  return Number.isFinite(rawZoom)
    ? normalizeOfficeZoom(Math.floor(rawZoom))
    : undefined;
}

/** 观察可视区域与内容尺寸，并持续维护自适应缩放比例。 */
export function useOfficeFitZoom({
  containerRef,
  mode,
  zoom,
  previewKind,
  sessionKey,
  activeKey,
  onZoom,
}: UseOfficeFitZoomOptions) {
  useEffect(() => {
    const container = containerRef.current;
    if (!container || !previewKind || mode === 'percentage') return;

    let animationFrame: number | undefined;
    let observedViewport: HTMLElement | undefined;
    let observedTarget: HTMLElement | undefined;
    let resizeObserver: ResizeObserver | undefined;

    const measure = () => {
      animationFrame = undefined;
      const nextZoom = measureFitZoom(container, mode, previewKind);
      if (nextZoom !== undefined && nextZoom !== zoom) onZoom(nextZoom);
    };
    const scheduleMeasure = () => {
      if (animationFrame !== undefined) return;
      animationFrame = window.requestAnimationFrame(measure);
    };
    resizeObserver =
      typeof ResizeObserver === 'undefined'
        ? undefined
        : new ResizeObserver(scheduleMeasure);
    const observeCurrentElements = () => {
      const viewport = container.querySelector<HTMLElement>(
        '[data-office-fit-viewport="true"]',
      );
      const target = container.querySelector<HTMLElement>(
        '[data-office-fit-target="true"]',
      );
      if (viewport === observedViewport && target === observedTarget) {
        scheduleMeasure();
        return;
      }
      resizeObserver?.disconnect();
      observedViewport = viewport ?? undefined;
      observedTarget = target ?? undefined;
      if (observedViewport) resizeObserver?.observe(observedViewport);
      if (observedTarget) resizeObserver?.observe(observedTarget);
      scheduleMeasure();
    };

    const mutationObserver =
      typeof MutationObserver === 'undefined'
        ? undefined
        : new MutationObserver(observeCurrentElements);
    mutationObserver?.observe(container, { childList: true, subtree: true });
    window.addEventListener('resize', scheduleMeasure);
    observeCurrentElements();

    return () => {
      if (animationFrame !== undefined) {
        window.cancelAnimationFrame(animationFrame);
      }
      resizeObserver?.disconnect();
      mutationObserver?.disconnect();
      window.removeEventListener('resize', scheduleMeasure);
    };
  }, [activeKey, containerRef, mode, onZoom, previewKind, sessionKey, zoom]);
}
