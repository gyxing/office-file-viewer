import type { KeyboardEvent, MouseEvent as ReactMouseEvent } from 'react';
import { useCallback, useEffect, useRef, useState } from 'react';

const DEFAULT_NOTES_HEIGHT = 220;
const MIN_NOTES_HEIGHT = 120;
const MAX_WORKSPACE_RATIO = 0.65;
const KEYBOARD_RESIZE_STEP = 20;

/** 管理备注面板的显式分隔条拖拽，避免依赖浏览器不一致的原生 resize 手柄。 */
export function useSpeakerNotesResize() {
  const panelRef = useRef<HTMLElement | null>(null);
  const cleanupDragRef = useRef<() => void>();
  const [height, setHeight] = useState(DEFAULT_NOTES_HEIGHT);
  const [maxHeight, setMaxHeight] = useState(DEFAULT_NOTES_HEIGHT);

  const readMaxHeight = useCallback(() => {
    const workspaceHeight =
      panelRef.current?.parentElement?.getBoundingClientRect().height ??
      DEFAULT_NOTES_HEIGHT / MAX_WORKSPACE_RATIO;
    return Math.max(MIN_NOTES_HEIGHT, workspaceHeight * MAX_WORKSPACE_RATIO);
  }, []);
  const clampHeight = useCallback(
    (value: number) =>
      Math.min(readMaxHeight(), Math.max(MIN_NOTES_HEIGHT, value)),
    [readMaxHeight],
  );

  useEffect(() => {
    const workspace = panelRef.current?.parentElement;
    const syncMaxHeight = () => {
      const nextMaxHeight = readMaxHeight();
      setMaxHeight(nextMaxHeight);
      setHeight((value) =>
        Math.min(nextMaxHeight, Math.max(MIN_NOTES_HEIGHT, value)),
      );
    };
    syncMaxHeight();

    if (!workspace || typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(syncMaxHeight);
    observer.observe(workspace);
    return () => observer.disconnect();
  }, [readMaxHeight]);

  const handleMouseDown = useCallback(
    (event: ReactMouseEvent<HTMLDivElement>) => {
      cleanupDragRef.current?.();
      const startHeight = height;
      const startY = event.clientY;
      const handleMouseMove = (moveEvent: globalThis.MouseEvent) => {
        setHeight(clampHeight(startHeight + startY - moveEvent.clientY));
      };
      const cleanup = () => {
        window.removeEventListener('mousemove', handleMouseMove);
        window.removeEventListener('mouseup', cleanup);
        cleanupDragRef.current = undefined;
      };
      cleanupDragRef.current = cleanup;
      window.addEventListener('mousemove', handleMouseMove);
      window.addEventListener('mouseup', cleanup);
      event.preventDefault();
    },
    [clampHeight, height],
  );

  useEffect(
    () => () => {
      cleanupDragRef.current?.();
    },
    [],
  );

  const handleKeyDown = useCallback(
    (event: KeyboardEvent<HTMLDivElement>) => {
      if (event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return;
      event.preventDefault();
      setHeight((value) =>
        clampHeight(
          value +
            (event.key === 'ArrowUp'
              ? KEYBOARD_RESIZE_STEP
              : -KEYBOARD_RESIZE_STEP),
        ),
      );
    },
    [clampHeight],
  );

  return {
    height,
    maxHeight,
    panelRef,
    handleKeyDown,
    handleMouseDown,
  };
}
