import type { KeyboardEvent } from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { createOfficeSearchSession } from '../../services/search/createOfficeSearchSession';
import { normalizeSearchText } from '../../services/search/normalizeSearchText';
import type {
  OfficeFileViewerSearchOptions,
  OfficeSearchProvider,
  OfficeSearchQuery,
  OfficeSearchResult,
  OfficeSearchTarget,
} from '../../services/search/types';

/** 搜索输入停止后再启动大文件扫描的短暂缓冲时间。 */
const OFFICE_SEARCH_INPUT_DELAY_MS = 120;

/** 格式查看器完成实际定位后返回是否成功。 */
export type OfficeSearchNavigator = (
  target: OfficeSearchTarget,
) => boolean | Promise<boolean>;

/** 搜索侧栏可读取的不可变状态。 */
export type OfficeSearchControllerState = Readonly<{
  query: string;
  matchCase: boolean;
  wholeWord: boolean;
  results: readonly OfficeSearchResult[];
  currentIndex: number;
  scanned: number;
  total: number;
  searching: boolean;
  complete: boolean;
  error?: string;
}>;

/** 搜索侧栏和格式查看器共享的稳定操作集合。 */
export type OfficeSearchControllerActions = Readonly<{
  setQuery(value: string): void;
  setMatchCase(value: boolean): void;
  setWholeWord(value: boolean): void;
  selectResult(index: number): Promise<boolean>;
  previousResult(): Promise<boolean>;
  nextResult(): Promise<boolean>;
  registerProvider(provider: OfficeSearchProvider): () => void;
  registerNavigator(
    kind: OfficeSearchTarget['kind'],
    navigator: OfficeSearchNavigator,
  ): () => void;
}>;

/** 搜索控制器对组合层公开的状态和动作。 */
export type OfficeSearchController = Readonly<{
  state: OfficeSearchControllerState;
  actions: OfficeSearchControllerActions;
  /** viewer 根节点统一处理 Ctrl+F 和 Esc。 */
  handleViewerKeyDown(event: KeyboardEvent<HTMLElement>): void;
}>;

/** 管理当前文件的搜索提供器、查询生命周期、结果选择和格式导航。 */
export function useOfficeSearchController({
  enabled,
  visible,
  sessionKey,
  options,
  hasDocument,
  onOpen,
  onClose,
}: {
  /** 当前实例是否启用搜索。 */
  enabled: boolean;
  /** 搜索侧栏当前是否展开。 */
  visible: boolean;
  /** 当前解析会话的稳定标识。 */
  sessionKey?: string;
  /** 初始搜索选项。 */
  options: OfficeFileViewerSearchOptions;
  /** 当前是否已有可搜索文档。 */
  hasDocument: boolean;
  /** 请求打开搜索侧栏。 */
  onOpen(): void;
  /** 请求关闭搜索侧栏。 */
  onClose(): void;
}): OfficeSearchController {
  const [provider, setProvider] = useState<OfficeSearchProvider>();
  const providerRef = useRef<OfficeSearchProvider>();
  const navigatorsRef = useRef(
    new Map<OfficeSearchTarget['kind'], OfficeSearchNavigator>(),
  );
  const [query, setQuery] = useState('');
  const [matchCase, setMatchCase] = useState(Boolean(options.matchCase));
  const [wholeWord, setWholeWord] = useState(Boolean(options.wholeWord));
  const [results, setResults] = useState<readonly OfficeSearchResult[]>([]);
  const resultsRef = useRef<readonly OfficeSearchResult[]>(results);
  const [currentIndex, setCurrentIndex] = useState(-1);
  const [scanned, setScanned] = useState(0);
  const [total, setTotal] = useState(0);
  const [searching, setSearching] = useState(false);
  const [complete, setComplete] = useState(true);
  const [error, setError] = useState<string>();
  const navigationGenerationRef = useRef(0);
  resultsRef.current = results;

  const session = useMemo(
    () => (provider ? createOfficeSearchSession(provider) : undefined),
    [provider, sessionKey],
  );

  useEffect(() => () => session?.dispose(), [session]);

  useEffect(() => {
    setQuery('');
    setMatchCase(Boolean(options.matchCase));
    setWholeWord(Boolean(options.wholeWord));
    setResults([]);
    setCurrentIndex(-1);
    setScanned(0);
    setTotal(0);
    setSearching(false);
    setComplete(true);
    setError(undefined);
    navigationGenerationRef.current += 1;
  }, [options.matchCase, options.wholeWord, sessionKey]);

  useEffect(() => {
    if (!enabled || !visible || !session) {
      session?.cancel();
      setSearching(false);
      return;
    }
    const searchQuery: OfficeSearchQuery = {
      text: query,
      matchCase,
      wholeWord,
    };
    if (!normalizeSearchText(query, matchCase).text) {
      session.cancel();
      setResults([]);
      setCurrentIndex(-1);
      setScanned(0);
      setTotal(0);
      setSearching(false);
      setComplete(true);
      setError(undefined);
      return;
    }

    let active = true;
    setResults([]);
    setCurrentIndex(-1);
    setScanned(0);
    setTotal(0);
    setSearching(true);
    setComplete(false);
    setError(undefined);
    const timer = setTimeout(() => {
      void session
        .search(searchQuery, (batch) => {
          if (!active) return;
          setResults((current) => {
            return [...current, ...batch.items];
          });
          setScanned(batch.scanned);
          setTotal(batch.total);
          setComplete(batch.complete);
          setSearching(!batch.complete);
        })
        .catch((searchError: unknown) => {
          if (!active) return;
          setSearching(false);
          setComplete(true);
          setError(
            searchError instanceof Error
              ? searchError.message
              : String(searchError),
          );
        });
    }, OFFICE_SEARCH_INPUT_DELAY_MS);

    return () => {
      active = false;
      clearTimeout(timer);
      session.cancel();
    };
  }, [enabled, matchCase, query, session, visible, wholeWord]);

  const registerProvider = useCallback((nextProvider: OfficeSearchProvider) => {
    providerRef.current = nextProvider;
    setProvider(nextProvider);
    return () => {
      if (providerRef.current !== nextProvider) return;
      providerRef.current = undefined;
      setProvider(undefined);
    };
  }, []);

  const registerNavigator = useCallback(
    (kind: OfficeSearchTarget['kind'], navigator: OfficeSearchNavigator) => {
      navigatorsRef.current.set(kind, navigator);
      return () => {
        if (navigatorsRef.current.get(kind) === navigator) {
          navigatorsRef.current.delete(kind);
        }
      };
    },
    [],
  );

  const selectResult = useCallback(async (index: number) => {
    const availableResults = resultsRef.current;
    if (!availableResults.length) return false;
    const normalizedIndex =
      ((Math.trunc(index) % availableResults.length) +
        availableResults.length) %
      availableResults.length;
    const result = availableResults[normalizedIndex];
    const navigator = navigatorsRef.current.get(result.target.kind);
    if (!navigator) return false;
    const generation = ++navigationGenerationRef.current;
    const navigated = await navigator(result.target);
    if (!navigated || generation !== navigationGenerationRef.current)
      return false;
    setCurrentIndex(normalizedIndex);
    return true;
  }, []);

  const previousResult = useCallback(
    () =>
      selectResult(currentIndex <= 0 ? results.length - 1 : currentIndex - 1),
    [currentIndex, results.length, selectResult],
  );
  const nextResult = useCallback(
    () => selectResult(currentIndex < 0 ? 0 : currentIndex + 1),
    [currentIndex, selectResult],
  );

  const handleViewerKeyDown = useCallback(
    (event: KeyboardEvent<HTMLElement>) => {
      if (
        enabled &&
        hasDocument &&
        (event.ctrlKey || event.metaKey) &&
        event.key.toLowerCase() === 'f'
      ) {
        event.preventDefault();
        onOpen();
        return;
      }
      if (visible && event.key === 'Escape') {
        event.preventDefault();
        onClose();
      }
    },
    [enabled, hasDocument, onClose, onOpen, visible],
  );

  return {
    state: {
      query,
      matchCase,
      wholeWord,
      results,
      currentIndex,
      scanned,
      total,
      searching,
      complete,
      error,
    },
    actions: {
      setQuery,
      setMatchCase,
      setWholeWord,
      selectResult,
      previousResult,
      nextResult,
      registerProvider,
      registerNavigator,
    },
    handleViewerKeyDown,
  };
}
