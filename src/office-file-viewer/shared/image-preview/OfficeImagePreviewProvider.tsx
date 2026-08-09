import type { ReactNode, RefObject } from 'react';
import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { resolveOfficeImagePreviewOptions } from './imagePreviewOptions';
import './index.less';
import { OfficeImageContextMenu } from './OfficeImageContextMenu';
import {
  OfficeImagePreviewContext,
  type OfficeImageContextMenuPoint,
} from './OfficeImagePreviewContext';
import { OfficeImagePreviewDialog } from './OfficeImagePreviewDialog';
import type {
  OfficeFileViewerImagePreviewConfig,
  OfficeImagePreviewTarget,
} from './types';

/** 图片预览 Provider 与当前预览器实例的连接参数。 */
type OfficeImagePreviewProviderProps = {
  /** 当前预览器包含的工具栏和格式渲染内容。 */
  children: ReactNode;
  /** 公共图片预览配置。 */
  config?: OfficeFileViewerImagePreviewConfig;
  /** 文件会话变化时用于清空临时图片交互状态的标识。 */
  sessionKey?: string;
  /** 图片弹层和菜单需要覆盖的预览器根元素。 */
  containerRef: RefObject<HTMLDivElement>;
};

/** 当前图片预览弹层关联的资源与触发元素。 */
type ActiveImagePreview = {
  /** 当前预览的图片资源。 */
  target: OfficeImagePreviewTarget;
  /** 关闭弹层后需要恢复焦点的触发元素。 */
  trigger: HTMLElement;
  /** 当前预览器所在的文档对象。 */
  ownerDocument: Document;
};

/** 当前图片右键菜单关联的位置和资源。 */
type ActiveImageContextMenu = ActiveImagePreview & {
  /** 菜单相对预览器左侧的位置。 */
  x: number;
  /** 菜单相对预览器顶部的位置。 */
  y: number;
};

/** 右键菜单固定宽度，用于在预览器边缘内修正横向位置。 */
const CONTEXT_MENU_WIDTH = 168;
/** 右键菜单含错误信息时预留的最大高度。 */
const CONTEXT_MENU_MAX_HEIGHT = 132;
/** 右键菜单与预览器边缘之间保留的最小间距。 */
const CONTEXT_MENU_INSET = 8;

/** 将右键菜单位置限制在当前预览器可视区域内。 */
function clampMenuPosition(value: number, available: number, size: number) {
  return Math.max(
    CONTEXT_MENU_INSET,
    Math.min(value, Math.max(CONTEXT_MENU_INSET, available - size)),
  );
}

/** 为一个 OfficeFileViewer 实例提供单例图片菜单和预览弹层。 */
export function OfficeImagePreviewProvider({
  children,
  config,
  sessionKey,
  containerRef,
}: OfficeImagePreviewProviderProps) {
  const options = useMemo(
    () => resolveOfficeImagePreviewOptions(config),
    [config],
  );
  const [preview, setPreview] = useState<ActiveImagePreview>();
  const [contextMenu, setContextMenu] = useState<ActiveImageContextMenu>();

  useEffect(() => {
    setPreview(undefined);
    setContextMenu(undefined);
  }, [options.enabled, sessionKey]);
  useEffect(() => {
    if (!options.contextMenu) setContextMenu(undefined);
  }, [options.contextMenu]);
  const openPreview = useCallback(
    (target: OfficeImagePreviewTarget, trigger: HTMLElement) => {
      if (!options.enabled) return;
      setContextMenu(undefined);
      setPreview({ target, trigger, ownerDocument: trigger.ownerDocument });
    },
    [options.enabled],
  );
  const openContextMenu = useCallback(
    (
      target: OfficeImagePreviewTarget,
      point: OfficeImageContextMenuPoint,
      trigger: HTMLElement,
    ) => {
      const container = containerRef.current;
      if (!container || !options.enabled || !options.contextMenu) return;
      const rect = container.getBoundingClientRect();
      setContextMenu({
        target,
        trigger,
        ownerDocument: trigger.ownerDocument,
        x: clampMenuPosition(
          point.clientX - rect.left,
          rect.width,
          CONTEXT_MENU_WIDTH,
        ),
        y: clampMenuPosition(
          point.clientY - rect.top,
          rect.height,
          CONTEXT_MENU_MAX_HEIGHT,
        ),
      });
    },
    [containerRef, options.contextMenu, options.enabled],
  );
  const closeContextMenu = useCallback(() => setContextMenu(undefined), []);
  const closePreview = useCallback(() => {
    setPreview((current) => {
      const focusTarget =
        current?.trigger && current.ownerDocument.contains(current.trigger)
          ? current.trigger
          : containerRef.current;
      current?.ownerDocument.defaultView?.requestAnimationFrame(() => {
        try {
          focusTarget?.focus({ preventScroll: true });
        } catch {
          focusTarget?.focus();
        }
      });
      return undefined;
    });
  }, [containerRef]);
  const previewContextMenuTarget = useCallback(() => {
    setContextMenu((current) => {
      if (current) {
        setPreview({
          target: current.target,
          trigger: current.trigger,
          ownerDocument: current.ownerDocument,
        });
      }
      return undefined;
    });
  }, []);
  const contextValue = useMemo(
    () => ({ options, openPreview, openContextMenu }),
    [openContextMenu, openPreview, options],
  );
  return (
    <OfficeImagePreviewContext.Provider value={contextValue}>
      {children}
      {contextMenu ? (
        <OfficeImageContextMenu
          target={contextMenu.target}
          x={contextMenu.x}
          y={contextMenu.y}
          download={options.download}
          ownerDocument={contextMenu.ownerDocument}
          onPreview={previewContextMenuTarget}
          onClose={closeContextMenu}
        />
      ) : null}
      {preview ? (
        <OfficeImagePreviewDialog
          key={`${sessionKey ?? 'session'}:${preview.target.id}`}
          target={preview.target}
          download={options.download}
          ownerDocument={preview.ownerDocument}
          onClose={closePreview}
        />
      ) : null}
    </OfficeImagePreviewContext.Provider>
  );
}
