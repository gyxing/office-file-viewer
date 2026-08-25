import { useEffect, useRef, useState } from 'react';

export type SiteCopyState = 'idle' | 'copied' | 'failed';

/** 优先使用 Clipboard API，并为权限受限或较旧的浏览器提供回退。 */
async function writeTextToClipboard(value: string): Promise<void> {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(value);
    return;
  }

  const textarea = document.createElement('textarea');
  textarea.value = value;
  textarea.setAttribute('readonly', '');
  textarea.style.position = 'fixed';
  textarea.style.opacity = '0';
  document.body.appendChild(textarea);
  textarea.select();

  const copied = document.execCommand('copy');
  textarea.remove();
  if (!copied) {
    throw new Error('Clipboard API is unavailable');
  }
}

/** 管理复制结果反馈，并在组件卸载时清理延迟任务。 */
export function useSiteCopyFeedback() {
  const [state, setState] = useState<SiteCopyState>('idle');
  const resetTimerRef = useRef<number>();

  useEffect(
    () => () => {
      if (resetTimerRef.current !== undefined) {
        window.clearTimeout(resetTimerRef.current);
      }
    },
    [],
  );

  const copy = async (value: string) => {
    if (resetTimerRef.current !== undefined) {
      window.clearTimeout(resetTimerRef.current);
    }

    try {
      await writeTextToClipboard(value);
      setState('copied');
    } catch {
      setState('failed');
    }

    resetTimerRef.current = window.setTimeout(() => setState('idle'), 2400);
  };

  return { copy, state };
}

/** 根据复制状态选择当前操作文案。 */
export function getSiteCopyLabel(
  state: SiteCopyState,
  labels: { idle: string; copied: string; failed: string },
): string {
  if (state === 'copied') return labels.copied;
  if (state === 'failed') return labels.failed;
  return labels.idle;
}
