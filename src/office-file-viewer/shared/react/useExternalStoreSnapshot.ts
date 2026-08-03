import * as React from 'react';

/** 外部状态发生变化时触发的监听函数。 */
type StoreChangeListener = () => void;

/** React 18 外部状态订阅 Hook 的最小兼容签名。 */
type SyncExternalStoreHook = <TSnapshot>(
  subscribe: (listener: StoreChangeListener) => () => void,
  getSnapshot: () => TSnapshot,
  getServerSnapshot?: () => TSnapshot,
) => TSnapshot;

/** React 16.9/17 类型中不存在、运行时可能由 React 18+ 提供的能力。 */
type ReactWithExternalStore = typeof React & {
  /** React 18+ 提供的并发安全外部状态订阅 Hook。 */
  useSyncExternalStore?: SyncExternalStoreHook;
};

/** 组件内部订阅所需的最小外部状态协议。 */
type ExternalStore<TSnapshot> = {
  /** 返回当前不可变快照；数据未变化时必须保持对象引用稳定。 */
  getSnapshot(): TSnapshot;
  /** 监听快照变化，并返回取消订阅函数。 */
  subscribe(listener: StoreChangeListener): () => void;
};

/** React 16.9/17 fallback 保存的最近一次渲染快照。 */
type FallbackStoreState<TSnapshot> = {
  /** 最近一次提交给 React 渲染的快照。 */
  value: TSnapshot;
  /** 读取当前快照的最新函数。 */
  getSnapshot: () => TSnapshot;
};

/** SSR 环境延后执行布局逻辑，避免服务端 useLayoutEffect 警告。 */
const useIsomorphicLayoutEffect =
  typeof window === 'undefined' ? React.useEffect : React.useLayoutEffect;

/** 空数据源不需要清理资源时复用的取消订阅函数。 */
const EMPTY_UNSUBSCRIBE = () => undefined;

/** 仅通过属性探测访问 React 18 API，避免 React 16/17 加载命名导入失败。 */
const nativeUseSyncExternalStore = (React as ReactWithExternalStore)
  .useSyncExternalStore;

/** 检查提交后的快照是否已经落后于数据源。 */
function hasSnapshotChanged<TSnapshot>(state: FallbackStoreState<TSnapshot>) {
  try {
    return !Object.is(state.value, state.getSnapshot());
  } catch {
    return true;
  }
}

/** 在 React 16.9/17 中补齐项目所需的最小外部状态订阅语义。 */
function useSyncExternalStoreFallback<TSnapshot>(
  subscribe: (listener: StoreChangeListener) => () => void,
  getSnapshot: () => TSnapshot,
  getServerSnapshot?: () => TSnapshot,
) {
  const value =
    typeof window === 'undefined' && getServerSnapshot
      ? getServerSnapshot()
      : getSnapshot();
  const [{ state }, forceUpdate] = React.useState(() => ({
    state: { value, getSnapshot },
  }));

  useIsomorphicLayoutEffect(() => {
    state.value = value;
    state.getSnapshot = getSnapshot;
    if (hasSnapshotChanged(state)) forceUpdate({ state });
  }, [getSnapshot, state, value]);

  useIsomorphicLayoutEffect(() => {
    const handleStoreChange = () => {
      if (hasSnapshotChanged(state)) forceUpdate({ state });
    };
    const unsubscribe = subscribe(handleStoreChange);
    // 订阅建立后立即复查，覆盖渲染至提交之间发生的更新。
    handleStoreChange();
    return unsubscribe;
  }, [state, subscribe]);

  return value;
}

/** React 18+ 保留原生并发语义，旧版本仅在运行时缺少 API 时使用 fallback。 */
const useSyncExternalStoreCompat: SyncExternalStoreHook =
  nativeUseSyncExternalStore ?? useSyncExternalStoreFallback;

/** 订阅可选数据源，并为所有 Office 格式统一稳定函数引用和版本兼容逻辑。 */
export function useExternalStoreSnapshot<TSnapshot>(
  source: ExternalStore<TSnapshot>,
): TSnapshot;
export function useExternalStoreSnapshot<TSnapshot>(
  source: ExternalStore<TSnapshot> | undefined,
  fallbackSnapshot: TSnapshot,
): TSnapshot;
export function useExternalStoreSnapshot<TSnapshot>(
  source: ExternalStore<TSnapshot> | undefined,
): TSnapshot | undefined;
export function useExternalStoreSnapshot<TSnapshot>(
  source: ExternalStore<TSnapshot> | undefined,
  fallbackSnapshot?: TSnapshot,
) {
  const subscribe = React.useCallback(
    (listener: StoreChangeListener) =>
      source === undefined ? EMPTY_UNSUBSCRIBE : source.subscribe(listener),
    [source],
  );
  const getSnapshot = React.useCallback(
    () => (source === undefined ? fallbackSnapshot : source.getSnapshot()),
    [fallbackSnapshot, source],
  );

  return useSyncExternalStoreCompat(subscribe, getSnapshot, getSnapshot);
}
