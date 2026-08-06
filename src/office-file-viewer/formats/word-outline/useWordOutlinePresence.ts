import { useEffect, useState } from 'react';

/**
 * 首次展开后保留当前文档的大纲实例，兼顾显隐动画和重复展开性能。
 */
export function useWordOutlinePresence(
  visible: boolean,
  documentSessionId: string,
) {
  const [activatedSessionId, setActivatedSessionId] = useState<
    string | undefined
  >(() => (visible ? documentSessionId : undefined));

  useEffect(() => {
    if (!visible) return;
    setActivatedSessionId((current) =>
      current === documentSessionId ? current : documentSessionId,
    );
  }, [documentSessionId, visible]);

  return visible || activatedSessionId === documentSessionId;
}
