import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  OFFICE_DEFAULT_ZOOM,
  OFFICE_MAX_ZOOM,
  OFFICE_MIN_ZOOM,
  OFFICE_ZOOM_STEP,
} from '../constants';
import { normalizeOfficeZoom } from '../normalizeOfficeZoom';
import type {
  OfficeViewerLayoutContentScaling,
  OfficeViewerLayoutContextValue,
} from './types';

/** 可复用外壳控制器输入。 */
type UseOfficeViewerLayoutControllerOptions = {
  /** 非受控缩放初始值。 */
  defaultZoom?: number;
  /** 受控缩放值。 */
  zoom?: number;
  /** 用户请求改变缩放时触发。 */
  onZoomChange?: (zoom: number) => void;
  /** 浏览器全屏状态改变时触发。 */
  onFullscreenChange?: (fullscreen: boolean) => void;
  /** 请求全屏失败时触发。 */
  onFullscreenError?: (error: Error) => void;
  /** 当前内容缩放职责。 */
  contentScaling: OfficeViewerLayoutContentScaling;
};

/** 将未知异常转换成宿主可稳定处理的 Error。 */
function normalizeFullscreenError(error: unknown): Error {
  return error instanceof Error ? error : new Error(String(error));
}

/** 管理复用外壳的缩放、全屏状态及公共上下文契约。 */
export function useOfficeViewerLayoutController({
  defaultZoom = OFFICE_DEFAULT_ZOOM,
  zoom,
  onZoomChange,
  onFullscreenChange,
  onFullscreenError,
  contentScaling,
}: UseOfficeViewerLayoutControllerOptions): OfficeViewerLayoutContextValue {
  const viewerRef = useRef<HTMLDivElement>(null);
  const [internalZoom, setInternalZoom] = useState(() =>
    normalizeOfficeZoom(defaultZoom),
  );
  const [isFullscreen, setIsFullscreen] = useState(false);
  const fullscreenRef = useRef(false);
  const resolvedZoom =
    zoom === undefined ? internalZoom : normalizeOfficeZoom(zoom);
  const fullscreenSupported =
    typeof document !== 'undefined' &&
    typeof document.documentElement.requestFullscreen === 'function';

  const changeZoom = useCallback(
    (nextZoom: number) => {
      const normalizedZoom = normalizeOfficeZoom(nextZoom, resolvedZoom);
      if (normalizedZoom === resolvedZoom) return;
      if (zoom === undefined) setInternalZoom(normalizedZoom);
      onZoomChange?.(normalizedZoom);
    },
    [onZoomChange, resolvedZoom, zoom],
  );
  const zoomOut = useCallback(() => {
    changeZoom(Math.max(OFFICE_MIN_ZOOM, resolvedZoom - OFFICE_ZOOM_STEP));
  }, [changeZoom, resolvedZoom]);
  const zoomIn = useCallback(() => {
    changeZoom(Math.min(OFFICE_MAX_ZOOM, resolvedZoom + OFFICE_ZOOM_STEP));
  }, [changeZoom, resolvedZoom]);

  useEffect(() => {
    if (typeof document === 'undefined') return undefined;
    const handleFullscreenChange = () => {
      const nextFullscreen = document.fullscreenElement === viewerRef.current;
      if (fullscreenRef.current === nextFullscreen) return;
      fullscreenRef.current = nextFullscreen;
      setIsFullscreen(nextFullscreen);
      onFullscreenChange?.(nextFullscreen);
    };
    document.addEventListener('fullscreenchange', handleFullscreenChange);
    return () => {
      document.removeEventListener('fullscreenchange', handleFullscreenChange);
    };
  }, [onFullscreenChange]);

  const toggleFullscreen = useCallback(() => {
    const viewer = viewerRef.current;
    if (!viewer || !fullscreenSupported || typeof document === 'undefined') {
      onFullscreenError?.(new Error('当前浏览器不支持标准全屏 API。'));
      return;
    }
    void (async () => {
      try {
        if (document.fullscreenElement === viewer) {
          await document.exitFullscreen();
        } else {
          await viewer.requestFullscreen();
        }
      } catch (error) {
        onFullscreenError?.(normalizeFullscreenError(error));
      }
    })();
  }, [fullscreenSupported, onFullscreenError]);

  const state = useMemo(
    () => ({ zoom: resolvedZoom, isFullscreen }),
    [isFullscreen, resolvedZoom],
  );
  const actions = useMemo(
    () => ({ changeZoom, zoomOut, zoomIn, toggleFullscreen }),
    [changeZoom, toggleFullscreen, zoomIn, zoomOut],
  );
  const meta = useMemo(
    () => ({ viewerRef, fullscreenSupported, contentScaling }),
    [contentScaling, fullscreenSupported],
  );

  return useMemo(() => ({ state, actions, meta }), [actions, meta, state]);
}
