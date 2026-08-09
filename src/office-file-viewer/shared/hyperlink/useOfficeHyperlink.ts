import type { KeyboardEvent, MouseEvent, PointerEvent } from 'react';
import { useCallback, useMemo, useRef } from 'react';
import { useOfficeHyperlinkContext } from './OfficeHyperlinkContext';
import type { OfficeHyperlink, OfficeHyperlinkSource } from './types';

/** 格式渲染节点绑定共享链接交互时使用的参数。 */
export type UseOfficeHyperlinkOptions = {
  /** 当前节点关联的标准链接；省略时不增加交互。 */
  hyperlink?: OfficeHyperlink;
  /** 当前节点的稳定来源身份。 */
  source: OfficeHyperlinkSource;
  /** 缩略图等只读副本可以显式关闭交互。 */
  interactive?: boolean;
};

/** 可直接合并到普通 HTML 元素的链接交互属性。 */
export type OfficeHyperlinkInteractionProps<TElement extends HTMLElement> = {
  role?: 'link';
  tabIndex?: number;
  title?: string;
  'data-office-hyperlink'?: 'true';
  onClick?: (event: MouseEvent<TElement>) => void;
  onContextMenu?: (event: MouseEvent<TElement>) => void;
  onKeyDown?: (event: KeyboardEvent<TElement>) => void;
  onPointerDown?: (event: PointerEvent<TElement>) => void;
};

/** 为任意格式渲染节点补充统一的链接激活规则。 */
export function useOfficeHyperlink<TElement extends HTMLElement>({
  hyperlink,
  source,
  interactive = true,
}: UseOfficeHyperlinkOptions): OfficeHyperlinkInteractionProps<TElement> {
  const context = useOfficeHyperlinkContext();
  const touchPointerRef = useRef(false);
  const enabled = Boolean(context?.enabled && hyperlink && interactive);

  const handlePointerDown = useCallback((event: PointerEvent<TElement>) => {
    touchPointerRef.current = event.pointerType === 'touch';
  }, []);
  const handleClick = useCallback(
    (event: MouseEvent<TElement>) => {
      if (!enabled || !context || !hyperlink || event.button !== 0) return;
      if (touchPointerRef.current) {
        touchPointerRef.current = false;
        event.preventDefault();
        event.stopPropagation();
        context.activate({ hyperlink, source, mode: 'touch' });
        return;
      }
      const modifierPressed =
        context.modifierLabel === 'Command' ? event.metaKey : event.ctrlKey;
      if (!modifierPressed) return;
      event.preventDefault();
      event.stopPropagation();
      context.activate({ hyperlink, source, mode: 'mouse' });
    },
    [context, enabled, hyperlink, source],
  );
  const handleContextMenu = useCallback(
    (event: MouseEvent<TElement>) => {
      if (!enabled || !context || !hyperlink || source.type === 'image') {
        return;
      }
      const opened = context.openContextMenu(
        { hyperlink, source, mode: 'context-menu' },
        { clientX: event.clientX, clientY: event.clientY },
        event.currentTarget,
      );
      if (!opened) return;
      event.preventDefault();
      event.stopPropagation();
      event.currentTarget.focus();
    },
    [context, enabled, hyperlink, source],
  );
  const handleKeyDown = useCallback(
    (event: KeyboardEvent<TElement>) => {
      if (!enabled || !context || !hyperlink) return;
      if (event.key === 'Enter') {
        event.preventDefault();
        event.stopPropagation();
        context.activate({ hyperlink, source, mode: 'keyboard' });
        return;
      }
      if (
        source.type === 'image' ||
        (event.key !== 'ContextMenu' &&
          !(event.shiftKey && event.key === 'F10'))
      ) {
        return;
      }
      const rect = event.currentTarget.getBoundingClientRect();
      const opened = context.openContextMenu(
        { hyperlink, source, mode: 'context-menu' },
        {
          clientX: rect.left + Math.min(rect.width, 24),
          clientY: rect.top + Math.min(rect.height, 24),
        },
        event.currentTarget,
      );
      if (!opened) return;
      event.preventDefault();
      event.stopPropagation();
    },
    [context, enabled, hyperlink, source],
  );

  return useMemo(() => {
    if (!enabled || !context || !hyperlink) return {};
    const title = hyperlink.screenTip
      ? `${hyperlink.screenTip}\n${context.activationHint}`
      : context.activationHint;
    return {
      role: 'link',
      tabIndex: 0,
      title,
      'data-office-hyperlink': 'true',
      onClick: handleClick,
      onContextMenu: handleContextMenu,
      onKeyDown: handleKeyDown,
      onPointerDown: handlePointerDown,
    };
  }, [
    context,
    enabled,
    handleClick,
    handleContextMenu,
    handleKeyDown,
    handlePointerDown,
    hyperlink,
  ]);
}
