import { useEffect, useRef, useState } from 'react';
import { useOfficeResourceStore } from './OfficeResourceContext';
import type { OfficeResourceSource } from './types';

/** 资源引用转换为浏览器地址时的加载状态。 */
type OfficeResourceUrlState = {
  /** 资源访问地址。 */
  url?: string;
  /** 加载状态相关文案。 */
  loading: boolean;
  /** 当前操作产生的错误；未提供表示没有错误。 */
  error?: Error;
};

/** 获取资源 URL，并在依赖变化或卸载时释放对应引用。 */
export function useOfficeResourceUrl(
  source: OfficeResourceSource | undefined,
): OfficeResourceUrlState {
  const store = useOfficeResourceStore();
  const generationRef = useRef(0);
  const [state, setState] = useState<OfficeResourceUrlState>({
    url: source?.kind === 'url' ? source.url : undefined,
    loading: source?.kind === 'lazy',
  });

  useEffect(() => {
    const generation = ++generationRef.current;
    if (!source) {
      setState({ loading: false });
      return;
    }
    if (source.kind === 'url') {
      setState({ url: source.url, loading: false });
      return;
    }
    if (!store) {
      setState({
        loading: false,
        error: new Error('当前预览树未提供 OfficeResourceStore'),
      });
      return;
    }

    const controller = new AbortController();
    let acquired = false;
    setState({ loading: true });
    void store.acquire(source, controller.signal).then(
      (url) => {
        acquired = true;
        if (generation === generationRef.current) {
          setState({ url, loading: false });
        }
      },
      (error) => {
        if (generation !== generationRef.current || controller.signal.aborted) {
          return;
        }
        setState({
          loading: false,
          error: error instanceof Error ? error : new Error('文档资源加载失败'),
        });
      },
    );

    return () => {
      controller.abort();
      if (acquired) store.release(source);
    };
  }, [source, store]);

  return state;
}
