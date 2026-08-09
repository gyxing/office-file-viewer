import type { ReactNode, RefObject } from 'react';
import React, {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type { PreviewKind } from '../../services/preview';
import type { OfficeFileViewerWarning } from '../../services/previewWarnings';
import { classifyOfficeHyperlink } from './classifyOfficeHyperlink';
import './index.less';
import {
  OfficeHyperlinkContext,
  type OfficeHyperlinkContextMenuPoint,
  type OfficeHyperlinkContextValue,
} from './OfficeHyperlinkContext';
import { OfficeHyperlinkContextMenu } from './OfficeHyperlinkContextMenu';
import type {
  OfficeHyperlinkActivateEvent,
  OfficeHyperlinkActivationRequest,
  OfficeHyperlinkNavigator,
} from './types';

/** 单实例超链接 Provider 的公共连接参数。 */
type OfficeHyperlinkProviderProps = {
  /** 当前预览器包含的格式渲染内容。 */
  children: ReactNode;
  /** 链接菜单需要覆盖并据此修正坐标的预览器根元素。 */
  containerRef: RefObject<HTMLDivElement>;
  /** 是否启用文档超链接交互。 */
  enabled: boolean;
  /** 当前正在预览的文件。 */
  file?: File;
  /** 当前文件的具体预览格式。 */
  previewKind?: PreviewKind;
  /** 远程文件的可靠来源 URL，用于解析相对地址。 */
  sourceUrl?: string;
  /** 切换文件时用于清空触屏确认状态的键。 */
  sessionKey?: string;
  /** 链接被有效激活时通知宿主。 */
  onActivate?: (event: OfficeHyperlinkActivateEvent) => void;
  /** 链接导航发生非致命降级时通知宿主。 */
  onWarning?: (warning: OfficeFileViewerWarning, file: File) => void;
};

/** 触屏二次确认的有效时间，避免很久后的点击被误认为确认。 */
const TOUCH_CONFIRM_TIMEOUT = 4000;
/** 超链接右键菜单的固定宽度。 */
const CONTEXT_MENU_WIDTH = 184;
/** 两项链接操作及错误信息可能占用的最大高度。 */
const CONTEXT_MENU_MAX_HEIGHT = 124;
/** 右键菜单与预览器边缘之间保留的最小间距。 */
const CONTEXT_MENU_INSET = 8;

/** 当前超链接右键菜单关联的链接、位置和触发元素。 */
type ActiveHyperlinkContextMenu = {
  /** 经过输入层封装的链接激活请求。 */
  request: OfficeHyperlinkActivationRequest;
  /** 当前菜单是否用于文档内部跳转。 */
  internal: boolean;
  /** 外部链接最终可复制的安全地址。 */
  copyTarget?: string;
  /** 菜单相对预览器左侧的位置。 */
  x: number;
  /** 菜单相对预览器顶部的位置。 */
  y: number;
  /** 当前预览器所在的文档对象。 */
  ownerDocument: Document;
};

/** 将右键菜单位置限制在当前预览器可视区域内。 */
function clampMenuPosition(value: number, available: number, size: number) {
  return Math.max(
    CONTEXT_MENU_INSET,
    Math.min(value, Math.max(CONTEXT_MENU_INSET, available - size)),
  );
}

/** 判断当前浏览器是否使用 Command 作为主修饰键。 */
function usesCommandModifier() {
  if (typeof navigator === 'undefined') return false;
  return /mac|iphone|ipad|ipod/i.test(
    `${navigator.platform ?? ''} ${navigator.userAgent ?? ''}`,
  );
}

/** 安全调用宿主观察回调，避免异常反向中断查看器事件。 */
function notifyObserver<TArguments extends unknown[]>(
  observer: ((...args: TArguments) => void) | undefined,
  ...args: TArguments
) {
  if (!observer) return;
  try {
    observer(...args);
  } catch (error) {
    setTimeout(() => {
      throw error;
    }, 0);
  }
}

/** 构造带可变 defaultPrevented 状态的公共激活事件。 */
function createActivateEvent(
  request: OfficeHyperlinkActivationRequest,
  file: File,
  previewKind: PreviewKind,
): OfficeHyperlinkActivateEvent {
  let defaultPrevented = false;
  return {
    hyperlink: request.hyperlink,
    file,
    previewKind,
    sourceType: request.source.type,
    sourceId: request.source.id,
    get defaultPrevented() {
      return defaultPrevented;
    },
    preventDefault() {
      defaultPrevented = true;
    },
  };
}

/** 为一个 OfficeFileViewer 实例统一提供超链接激活与内部导航。 */
export function OfficeHyperlinkProvider({
  children,
  containerRef,
  enabled,
  file,
  previewKind,
  sourceUrl,
  sessionKey,
  onActivate,
  onWarning,
}: OfficeHyperlinkProviderProps) {
  const messages = useOfficeFileViewerMessages();
  const commandModifier = usesCommandModifier();
  const modifierLabel = commandModifier ? 'Command' : 'Ctrl';
  const navigatorsRef = useRef(
    new Map<
      'word' | 'spreadsheet' | 'presentation',
      OfficeHyperlinkNavigator
    >(),
  );
  const touchConfirmationRef = useRef<{
    key: string;
    expiresAt: number;
  }>();
  const [touchHintVisible, setTouchHintVisible] = useState(false);
  const [contextMenu, setContextMenu] = useState<ActiveHyperlinkContextMenu>();

  const clearTouchConfirmation = useCallback(() => {
    touchConfirmationRef.current = undefined;
    setTouchHintVisible(false);
  }, []);

  useEffect(clearTouchConfirmation, [
    clearTouchConfirmation,
    enabled,
    sessionKey,
  ]);
  useEffect(() => setContextMenu(undefined), [enabled, sessionKey]);
  useEffect(() => {
    if (!touchHintVisible) return;
    const timer = window.setTimeout(
      clearTouchConfirmation,
      TOUCH_CONFIRM_TIMEOUT,
    );
    const ownerDocument =
      file && typeof document !== 'undefined' ? document : undefined;
    const handlePointerDown = (event: PointerEvent) => {
      const target = event.target;
      if (
        target instanceof Element &&
        target.closest('[data-office-hyperlink]')
      ) {
        return;
      }
      clearTouchConfirmation();
    };
    ownerDocument?.addEventListener('pointerdown', handlePointerDown, true);
    return () => {
      window.clearTimeout(timer);
      ownerDocument?.removeEventListener(
        'pointerdown',
        handlePointerDown,
        true,
      );
    };
  }, [clearTouchConfirmation, file, touchHintVisible]);

  const reportWarning = useCallback(
    (code: string, message: string) => {
      if (!file || !previewKind) return;
      notifyObserver(
        onWarning,
        { code, message, previewKind, source: 'hyperlink' },
        file,
      );
    },
    [file, onWarning, previewKind],
  );

  const activate = useCallback(
    (request: OfficeHyperlinkActivationRequest) => {
      if (!enabled || !file || !previewKind) return;
      if (request.mode === 'touch') {
        const key = `${request.source.type}:${request.source.id}`;
        const current = touchConfirmationRef.current;
        if (!current || current.key !== key || current.expiresAt < Date.now()) {
          touchConfirmationRef.current = {
            key,
            expiresAt: Date.now() + TOUCH_CONFIRM_TIMEOUT,
          };
          setTouchHintVisible(true);
          return;
        }
        clearTouchConfirmation();
      }

      const event = createActivateEvent(request, file, previewKind);
      notifyObserver(onActivate, event);
      if (event.defaultPrevented) return;
      const classified = classifyOfficeHyperlink(request.hyperlink, sourceUrl);

      if (classified.kind === 'host-only') return;
      if (classified.kind === 'blocked') {
        reportWarning('HYPERLINK_BLOCKED', messages.hyperlink.blocked);
        return;
      }
      if (classified.kind === 'new-tab') {
        const opened = window.open(
          classified.target,
          '_blank',
          'noopener,noreferrer',
        );
        if (opened) opened.opener = null;
        return;
      }
      if (classified.kind === 'system') {
        window.location.assign(classified.target);
        return;
      }

      const navigator = navigatorsRef.current.get(classified.target.family);
      if (!navigator) {
        reportWarning(
          'HYPERLINK_TARGET_NOT_FOUND',
          messages.hyperlink.targetNotFound,
        );
        return;
      }
      void Promise.resolve(navigator(classified.target))
        .then((navigated) => {
          if (!navigated) {
            reportWarning(
              'HYPERLINK_TARGET_NOT_FOUND',
              messages.hyperlink.targetNotFound,
            );
          }
        })
        .catch((error) => {
          if (error instanceof Error && error.name === 'AbortError') return;
          reportWarning(
            'HYPERLINK_TARGET_NOT_FOUND',
            messages.hyperlink.targetNotFound,
          );
        });
    },
    [
      clearTouchConfirmation,
      enabled,
      file,
      messages.hyperlink.blocked,
      messages.hyperlink.targetNotFound,
      onActivate,
      previewKind,
      reportWarning,
      sourceUrl,
    ],
  );

  const openContextMenu = useCallback(
    (
      request: OfficeHyperlinkActivationRequest,
      point: OfficeHyperlinkContextMenuPoint,
      trigger: HTMLElement,
    ) => {
      const container = containerRef.current;
      if (!enabled || !container || request.source.type === 'image') {
        return false;
      }
      const classified = classifyOfficeHyperlink(request.hyperlink, sourceUrl);
      // 本地路径、无可靠基准的相对路径和危险协议继续沿用原生右键行为。
      if (classified.kind === 'host-only' || classified.kind === 'blocked') {
        return false;
      }
      const rect = container.getBoundingClientRect();
      const copyTarget =
        classified.kind === 'new-tab' || classified.kind === 'system'
          ? classified.target
          : undefined;
      setContextMenu({
        request,
        internal: classified.kind === 'internal',
        copyTarget,
        ownerDocument: trigger.ownerDocument,
        x: clampMenuPosition(
          point.clientX - rect.left,
          rect.width,
          CONTEXT_MENU_WIDTH,
        ),
        y: clampMenuPosition(
          point.clientY - rect.top,
          rect.height,
          copyTarget ? CONTEXT_MENU_MAX_HEIGHT : 52,
        ),
      });
      clearTouchConfirmation();
      return true;
    },
    [clearTouchConfirmation, containerRef, enabled, sourceUrl],
  );
  const closeContextMenu = useCallback(() => setContextMenu(undefined), []);
  const activateContextMenuTarget = useCallback(() => {
    if (!contextMenu) return;
    const request = contextMenu.request;
    setContextMenu(undefined);
    activate({ ...request, mode: 'context-menu' });
  }, [activate, contextMenu]);

  const registerNavigator = useCallback<
    OfficeHyperlinkContextValue['registerNavigator']
  >((family, navigator) => {
    navigatorsRef.current.set(family, navigator);
    return () => {
      if (navigatorsRef.current.get(family) === navigator) {
        navigatorsRef.current.delete(family);
      }
    };
  }, []);
  const contextValue = useMemo<OfficeHyperlinkContextValue>(
    () => ({
      enabled,
      modifierLabel,
      activationHint: messages.hyperlink.activationHint(modifierLabel),
      activate,
      openContextMenu,
      registerNavigator,
    }),
    [
      activate,
      enabled,
      messages.hyperlink,
      modifierLabel,
      openContextMenu,
      registerNavigator,
    ],
  );

  return (
    <OfficeHyperlinkContext.Provider value={contextValue}>
      {children}
      {contextMenu ? (
        <OfficeHyperlinkContextMenu
          x={contextMenu.x}
          y={contextMenu.y}
          internal={contextMenu.internal}
          copyTarget={contextMenu.copyTarget}
          ownerDocument={contextMenu.ownerDocument}
          onOpen={activateContextMenuTarget}
          onClose={closeContextMenu}
        />
      ) : null}
      {touchHintVisible ? (
        <div
          className="office-file-hyperlink-hint"
          role="status"
          aria-live="polite"
        >
          {messages.hyperlink.touchConfirm}
        </div>
      ) : null}
    </OfficeHyperlinkContext.Provider>
  );
}
