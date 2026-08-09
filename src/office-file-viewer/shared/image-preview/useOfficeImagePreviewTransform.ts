import type {
  CSSProperties,
  PointerEvent as ReactPointerEvent,
  WheelEvent as ReactWheelEvent,
  SyntheticEvent,
} from 'react';
import { useEffect, useMemo, useRef, useState } from 'react';
import {
  calculateOfficeImageFitScale,
  normalizeOfficeImagePreviewZoom,
  OFFICE_IMAGE_PREVIEW_MAX_ZOOM,
  OFFICE_IMAGE_PREVIEW_MIN_ZOOM,
  OFFICE_IMAGE_PREVIEW_ZOOM_STEP,
} from './imagePreviewTransform';

/** 图片或预览区域的像素尺寸。 */
type ImagePreviewSize = {
  /** 宽度。 */
  width: number;
  /** 高度。 */
  height: number;
};

/** 图片相对预览区域中心的平移距离。 */
type ImagePreviewOffset = {
  /** 水平偏移量。 */
  x: number;
  /** 垂直偏移量。 */
  y: number;
};

/** 一次图片拖拽手势的起始状态。 */
type ImagePreviewDrag = {
  /** 当前捕获的指针标识。 */
  pointerId: number;
  /** 上一次指针事件的横坐标。 */
  clientX: number;
  /** 上一次指针事件的纵坐标。 */
  clientY: number;
};

/** 尚未测量时使用的空尺寸。 */
const EMPTY_SIZE: ImagePreviewSize = { width: 0, height: 0 };
/** 图片复位时使用的中心偏移。 */
const ZERO_OFFSET: ImagePreviewOffset = { x: 0, y: 0 };
/** 适应窗口时为图片四周预留的总间距。 */
const IMAGE_VIEWPORT_INSET = 48;

/** 管理图片预览的适应窗口、缩放、旋转和拖拽状态。 */
export function useOfficeImagePreviewTransform(targetId: string) {
  const viewportRef = useRef<HTMLDivElement>(null);
  const imageRef = useRef<HTMLImageElement>(null);
  const dragRef = useRef<ImagePreviewDrag>();
  const offsetRef = useRef<ImagePreviewOffset>(ZERO_OFFSET);
  const pendingOffsetRef = useRef<ImagePreviewOffset>();
  const animationFrameRef = useRef<number>();
  const [viewportSize, setViewportSize] = useState(EMPTY_SIZE);
  const [imageSize, setImageSize] = useState(EMPTY_SIZE);
  const [zoom, setZoom] = useState(1);
  const [rotation, setRotation] = useState(0);
  const [offset, setOffset] = useState(ZERO_OFFSET);
  const [dragging, setDragging] = useState(false);

  useEffect(() => {
    setImageSize(EMPTY_SIZE);
    setZoom(1);
    setRotation(0);
    offsetRef.current = ZERO_OFFSET;
    setOffset(ZERO_OFFSET);
  }, [targetId]);

  useEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return undefined;
    const measure = () =>
      setViewportSize({
        width: viewport.clientWidth,
        height: viewport.clientHeight,
      });
    measure();
    if (typeof ResizeObserver === 'undefined') {
      window.addEventListener('resize', measure);
      return () => window.removeEventListener('resize', measure);
    }
    const observer = new ResizeObserver(measure);
    observer.observe(viewport);
    return () => observer.disconnect();
  }, []);
  useEffect(
    () => () => {
      const frame = animationFrameRef.current;
      if (frame !== undefined) cancelAnimationFrame(frame);
    },
    [],
  );

  const scheduleOffset = (nextOffset: ImagePreviewOffset) => {
    offsetRef.current = nextOffset;
    pendingOffsetRef.current = nextOffset;
    if (animationFrameRef.current !== undefined) return;
    const view = viewportRef.current?.ownerDocument.defaultView;
    const requestFrame = view?.requestAnimationFrame ?? requestAnimationFrame;
    animationFrameRef.current = requestFrame(() => {
      animationFrameRef.current = undefined;
      const pendingOffset = pendingOffsetRef.current;
      if (pendingOffset) setOffset(pendingOffset);
    });
  };

  const handleImageLoad = (event: SyntheticEvent<HTMLImageElement>) => {
    const image = event.currentTarget;
    setImageSize({ width: image.naturalWidth, height: image.naturalHeight });
  };

  const handlePointerDown = (event: ReactPointerEvent<HTMLDivElement>) => {
    const image = imageRef.current;
    // 空白区域用于关闭预览，拖拽手势只能从图片本身开始。
    if (
      event.button !== 0 ||
      !imageSize.width ||
      !image ||
      event.target !== image
    ) {
      return;
    }
    image.setPointerCapture(event.pointerId);
    dragRef.current = {
      pointerId: event.pointerId,
      clientX: event.clientX,
      clientY: event.clientY,
    };
    setDragging(true);
  };

  const handlePointerMove = (event: ReactPointerEvent<HTMLDivElement>) => {
    const drag = dragRef.current;
    if (!drag || drag.pointerId !== event.pointerId) return;
    const nextOffset = {
      x: offsetRef.current.x + event.clientX - drag.clientX,
      y: offsetRef.current.y + event.clientY - drag.clientY,
    };
    drag.clientX = event.clientX;
    drag.clientY = event.clientY;
    scheduleOffset(nextOffset);
  };

  const finishDragging = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (dragRef.current?.pointerId !== event.pointerId) return;
    dragRef.current = undefined;
    const image = imageRef.current;
    if (image?.hasPointerCapture(event.pointerId)) {
      image.releasePointerCapture(event.pointerId);
    }
    setDragging(false);
  };
  const changeZoom = (delta: number) => {
    setZoom((current) => normalizeOfficeImagePreviewZoom(current + delta));
  };
  const handleWheel = (event: ReactWheelEvent<HTMLDivElement>) => {
    event.preventDefault();
    if (event.deltaY === 0) return;
    changeZoom(
      event.deltaY < 0
        ? OFFICE_IMAGE_PREVIEW_ZOOM_STEP
        : -OFFICE_IMAGE_PREVIEW_ZOOM_STEP,
    );
  };
  const rotateClockwise = () => {
    setRotation((current) => (current + 90) % 360);
    offsetRef.current = ZERO_OFFSET;
    setOffset(ZERO_OFFSET);
  };
  const reset = () => {
    setZoom(1);
    setRotation(0);
    offsetRef.current = ZERO_OFFSET;
    setOffset(ZERO_OFFSET);
  };
  const fitScale = calculateOfficeImageFitScale(
    imageSize.width,
    imageSize.height,
    Math.max(0, viewportSize.width - IMAGE_VIEWPORT_INSET),
    Math.max(0, viewportSize.height - IMAGE_VIEWPORT_INSET),
    rotation,
  );
  const imageStyle = useMemo<CSSProperties>(
    () => ({
      width: imageSize.width || undefined,
      height: imageSize.height || undefined,
      visibility: imageSize.width ? 'visible' : 'hidden',
      transform: `translate(-50%, -50%) translate3d(${offset.x}px, ${
        offset.y
      }px, 0) rotate(${rotation}deg) scale(${fitScale * zoom})`,
    }),
    [
      fitScale,
      imageSize.height,
      imageSize.width,
      offset.x,
      offset.y,
      rotation,
      zoom,
    ],
  );

  return {
    viewportRef,
    imageRef,
    imageStyle,
    zoomPercent: Math.round(zoom * 100),
    canZoomOut: zoom > OFFICE_IMAGE_PREVIEW_MIN_ZOOM,
    canZoomIn: zoom < OFFICE_IMAGE_PREVIEW_MAX_ZOOM,
    dragging,
    handleImageLoad,
    handlePointerDown,
    handlePointerMove,
    handlePointerUp: finishDragging,
    handlePointerCancel: finishDragging,
    handleWheel,
    zoomOut: () => changeZoom(-OFFICE_IMAGE_PREVIEW_ZOOM_STEP),
    zoomIn: () => changeZoom(OFFICE_IMAGE_PREVIEW_ZOOM_STEP),
    rotateClockwise,
    reset,
  };
}
