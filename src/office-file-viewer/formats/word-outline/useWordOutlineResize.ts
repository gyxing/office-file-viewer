import type {
  KeyboardEvent,
  PointerEvent as ReactPointerEvent,
  RefObject,
} from 'react';
import { useCallback, useEffect, useRef, useState } from 'react';

/** Word 大纲侧栏的默认宽度，单位为 CSS 像素。 */
const DEFAULT_OUTLINE_WIDTH = 260;
/** Word 大纲侧栏允许的最小宽度，单位为 CSS 像素。 */
const MIN_OUTLINE_WIDTH = 200;
/** Word 大纲侧栏允许的绝对最大宽度，单位为 CSS 像素。 */
const MAX_OUTLINE_WIDTH = 520;
/** 大纲侧栏最多占用预览工作区宽度的比例。 */
const MAX_WORKSPACE_WIDTH_RATIO = 0.45;
/** 键盘调整大纲侧栏宽度时的单次步长，单位为 CSS 像素。 */
const KEYBOARD_RESIZE_STEP = 20;
/** 拖拽宽度写入侧栏时使用的 CSS 自定义属性。 */
const OUTLINE_WIDTH_CSS_VARIABLE = '--office-file-word-outline-width';

type OutlineResizeDrag = {
  /** 当前由分隔条捕获的指针标识。 */
  pointerId: number;
  /** 开始拖拽时的水平坐标。 */
  startX: number;
  /** 开始拖拽时的侧栏宽度。 */
  startWidth: number;
};

/** 将侧栏宽度限制在当前工作区允许的范围内。 */
function clampOutlineWidth(value: number, maxWidth: number) {
  return Math.min(maxWidth, Math.max(MIN_OUTLINE_WIDTH, value));
}

/**
 * 管理 Word 大纲侧栏的横向缩放。
 *
 * 拖动期间直接按动画帧更新 CSS 变量，避免大纲树和正文被高频重新渲染。
 */
export function useWordOutlineResize(
  panelRef: RefObject<HTMLElement>,
  documentSessionId: string,
) {
  const handleRef = useRef<HTMLDivElement>(null);
  const widthRef = useRef(DEFAULT_OUTLINE_WIDTH);
  const maxWidthRef = useRef(MAX_OUTLINE_WIDTH);
  const dragRef = useRef<OutlineResizeDrag>();
  const pendingWidthRef = useRef<number>();
  const animationFrameRef = useRef<number>();
  const [width, setWidth] = useState(DEFAULT_OUTLINE_WIDTH);
  const [maxWidth, setMaxWidth] = useState(MAX_OUTLINE_WIDTH);

  const readMaxWidth = useCallback(() => {
    const workspaceWidth =
      panelRef.current?.parentElement?.getBoundingClientRect().width;
    if (!workspaceWidth || !Number.isFinite(workspaceWidth)) {
      return MAX_OUTLINE_WIDTH;
    }
    return Math.max(
      MIN_OUTLINE_WIDTH,
      Math.min(MAX_OUTLINE_WIDTH, workspaceWidth * MAX_WORKSPACE_WIDTH_RATIO),
    );
  }, [panelRef]);

  const applyWidth = useCallback(
    (value: number) => {
      const nextWidth = clampOutlineWidth(value, maxWidthRef.current);
      widthRef.current = nextWidth;
      panelRef.current?.style.setProperty(
        OUTLINE_WIDTH_CSS_VARIABLE,
        `${nextWidth}px`,
      );
      handleRef.current?.setAttribute(
        'aria-valuenow',
        `${Math.round(nextWidth)}`,
      );
      return nextWidth;
    },
    [panelRef],
  );

  const syncWidthRange = useCallback(() => {
    const nextMaxWidth = readMaxWidth();
    maxWidthRef.current = nextMaxWidth;
    handleRef.current?.setAttribute(
      'aria-valuemax',
      `${Math.round(nextMaxWidth)}`,
    );
    const nextWidth = applyWidth(widthRef.current);
    setMaxWidth((current) =>
      current === nextMaxWidth ? current : nextMaxWidth,
    );
    setWidth((current) => (current === nextWidth ? current : nextWidth));
    return nextWidth;
  }, [applyWidth, readMaxWidth]);

  const scheduleWidth = useCallback(
    (value: number) => {
      pendingWidthRef.current = value;
      if (animationFrameRef.current !== undefined) return;
      if (typeof requestAnimationFrame === 'undefined') {
        pendingWidthRef.current = undefined;
        applyWidth(value);
        return;
      }
      animationFrameRef.current = requestAnimationFrame(() => {
        animationFrameRef.current = undefined;
        const pendingWidth = pendingWidthRef.current;
        pendingWidthRef.current = undefined;
        if (pendingWidth !== undefined) applyWidth(pendingWidth);
      });
    },
    [applyWidth],
  );

  const flushScheduledWidth = useCallback(() => {
    if (
      animationFrameRef.current !== undefined &&
      typeof cancelAnimationFrame !== 'undefined'
    ) {
      cancelAnimationFrame(animationFrameRef.current);
    }
    animationFrameRef.current = undefined;
    const pendingWidth = pendingWidthRef.current;
    pendingWidthRef.current = undefined;
    return pendingWidth === undefined
      ? widthRef.current
      : applyWidth(pendingWidth);
  }, [applyWidth]);

  useEffect(() => {
    widthRef.current = DEFAULT_OUTLINE_WIDTH;
    const workspace = panelRef.current?.parentElement;
    syncWidthRange();

    const observer =
      workspace && typeof ResizeObserver !== 'undefined'
        ? new ResizeObserver(syncWidthRange)
        : undefined;
    if (observer && workspace) observer.observe(workspace);

    return () => {
      observer?.disconnect();
      panelRef.current?.removeAttribute('data-resizing');
      if (
        animationFrameRef.current !== undefined &&
        typeof cancelAnimationFrame !== 'undefined'
      ) {
        cancelAnimationFrame(animationFrameRef.current);
      }
      animationFrameRef.current = undefined;
      pendingWidthRef.current = undefined;
      const pointerId = dragRef.current?.pointerId;
      dragRef.current = undefined;
      if (
        pointerId !== undefined &&
        handleRef.current?.hasPointerCapture(pointerId)
      ) {
        handleRef.current.releasePointerCapture(pointerId);
      }
    };
  }, [documentSessionId, panelRef, syncWidthRange]);

  const handlePointerDown = useCallback(
    (event: ReactPointerEvent<HTMLDivElement>) => {
      if (!event.isPrimary || event.button !== 0) return;
      const startWidth = syncWidthRange();
      dragRef.current = {
        pointerId: event.pointerId,
        startX: event.clientX,
        startWidth,
      };
      panelRef.current?.setAttribute('data-resizing', 'true');
      event.currentTarget.setPointerCapture(event.pointerId);
      event.preventDefault();
    },
    [syncWidthRange],
  );

  const handlePointerMove = useCallback(
    (event: ReactPointerEvent<HTMLDivElement>) => {
      const drag = dragRef.current;
      if (!drag || drag.pointerId !== event.pointerId) return;
      scheduleWidth(drag.startWidth + event.clientX - drag.startX);
      event.preventDefault();
    },
    [scheduleWidth],
  );

  const handlePointerEnd = useCallback(
    (event: ReactPointerEvent<HTMLDivElement>) => {
      const drag = dragRef.current;
      if (!drag || drag.pointerId !== event.pointerId) return;
      const nextWidth = flushScheduledWidth();
      dragRef.current = undefined;
      panelRef.current?.removeAttribute('data-resizing');
      setWidth((current) => (current === nextWidth ? current : nextWidth));
      if (event.currentTarget.hasPointerCapture(event.pointerId)) {
        event.currentTarget.releasePointerCapture(event.pointerId);
      }
    },
    [flushScheduledWidth],
  );

  const handleKeyDown = useCallback(
    (event: KeyboardEvent<HTMLDivElement>) => {
      if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
      event.preventDefault();
      syncWidthRange();
      const nextWidth = applyWidth(
        widthRef.current +
          (event.key === 'ArrowRight'
            ? KEYBOARD_RESIZE_STEP
            : -KEYBOARD_RESIZE_STEP),
      );
      setWidth((current) => (current === nextWidth ? current : nextWidth));
    },
    [applyWidth, syncWidthRange],
  );

  return {
    width,
    maxWidth,
    minWidth: MIN_OUTLINE_WIDTH,
    handleRef,
    handleKeyDown,
    handlePointerDown,
    handlePointerMove,
    handlePointerEnd,
  };
}
