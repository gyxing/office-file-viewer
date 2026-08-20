import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { OfficeAnnotationSource } from '../../services/annotations/AnnotationSource';
import type {
  OfficeAnnotation,
  OfficeAnnotationSourceSnapshot,
  OfficeAnnotationTarget,
} from '../../services/annotations/types';
import type { WordRevisionRecord } from '../../services/word/review/types';

/** 尚无格式数据源时复用的稳定空快照。 */
const EMPTY_ANNOTATION_SNAPSHOT: OfficeAnnotationSourceSnapshot = {
  revision: 0,
  count: 0,
  revisionCount: 0,
  noteCount: 0,
  supportsRevisionModes: false,
};

/** 格式查看器完成批注目标定位后返回是否成功。 */
export type OfficeAnnotationNavigator = (
  annotation: OfficeAnnotation,
) => boolean | Promise<boolean>;

/** 审阅面板可读取的批注状态。 */
export type OfficeAnnotationControllerState = Readonly<{
  /** 当前格式数据源的轻量快照。 */
  snapshot: OfficeAnnotationSourceSnapshot;
  /** 当前活动批注的索引。 */
  activeIndex: number;
  /** 当前活动批注；尚未选择时为空。 */
  activeAnnotation?: OfficeAnnotation;
  /** 当前活动修订；尚未选择时为空。 */
  activeRevision?: WordRevisionRecord;
  /** 可见范围是否正在读取。 */
  loading: boolean;
  /** 最近一次读取或定位失败的说明。 */
  error?: string;
  /** 批注缓存变化标识，供窗口列表重新取值。 */
  cacheRevision: number;
}>;

/** 审阅面板和格式查看器共享的稳定操作集合。 */
export type OfficeAnnotationControllerActions = Readonly<{
  /** 注册当前格式的批注数据源。 */
  registerSource(source: OfficeAnnotationSource): () => void;
  /** 注册指定目标类别的精确导航能力。 */
  registerNavigator(
    kind: OfficeAnnotationTarget['kind'],
    navigator: OfficeAnnotationNavigator,
  ): () => void;
  /** 确保指定批注索引范围已经进入缓存。 */
  loadRange(start: number, end: number): Promise<void>;
  /** 读取缓存中的指定批注。 */
  getCached(index: number): OfficeAnnotation | undefined;
  /** 按稳定标识读取一组批注，但不改变当前活动项。 */
  loadAnnotationsByIds(
    ids: readonly string[],
  ): Promise<readonly OfficeAnnotation[]>;
  /** 选择指定索引并导航到源内容。 */
  selectIndex(index: number): Promise<boolean>;
  /** 通过稳定标识选择批注。 */
  selectId(id: string): Promise<boolean>;
  /** 选择上一条批注。 */
  previous(): Promise<boolean>;
  /** 选择下一条批注。 */
  next(): Promise<boolean>;
  /** 从正文激活修订，不把当前页面跳回该修订的首个片段。 */
  activateRevisionId(id: string): Promise<boolean>;
  /** 关闭正文中已经固定显示的修订提示框。 */
  clearActiveRevision(): void;
}>;

/** 组合层与格式渲染器共享的批注控制器。 */
export type OfficeAnnotationController = Readonly<{
  /** 当前审阅运行时启用的内容类别。 */
  options: Readonly<{
    showComments: boolean;
    showNotes: boolean;
  }>;
  state: OfficeAnnotationControllerState;
  actions: OfficeAnnotationControllerActions;
}>;

/** 判断异常是否由面板切换或滚动产生的主动取消。 */
function isAbortError(error: unknown) {
  return error instanceof Error && error.name === 'AbortError';
}

/** 把批注索引约束到当前数据源范围内。 */
function normalizeAnnotationIndex(index: number, count: number) {
  if (!count) return -1;
  return Math.min(count - 1, Math.max(0, Math.trunc(index)));
}

/** 管理批注数据源、窗口缓存、活动项和格式导航。 */
export function useOfficeAnnotationController({
  enabled,
  sessionKey,
  showComments,
  showNotes,
  onActivate,
}: {
  /** 当前实例是否启用审阅能力。 */
  enabled: boolean;
  /** 当前解析会话的稳定标识。 */
  sessionKey?: string;
  /** 是否在正文和面板中显示批注。 */
  showComments: boolean;
  /** 是否在审阅摘要中显示脚注、尾注和格式笔记。 */
  showNotes: boolean;
  /** 批注从正文或外部导航激活后请求打开审阅面板。 */
  onActivate?(): void;
}): OfficeAnnotationController {
  const [source, setSource] = useState<OfficeAnnotationSource>();
  const sourceRef = useRef<OfficeAnnotationSource>();
  const [snapshot, setSnapshot] = useState(EMPTY_ANNOTATION_SNAPSHOT);
  const snapshotRef = useRef(snapshot);
  const cacheRef = useRef(new Map<number, OfficeAnnotation>());
  const [cacheRevision, setCacheRevision] = useState(0);
  const [activeIndex, setActiveIndex] = useState(-1);
  const [activeAnnotation, setActiveAnnotation] = useState<OfficeAnnotation>();
  const [activeRevision, setActiveRevision] = useState<WordRevisionRecord>();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>();
  const readGenerationRef = useRef(0);
  const navigationGenerationRef = useRef(0);
  const revisionSelectionGenerationRef = useRef(0);
  const readControllerRef = useRef<AbortController>();
  const navigatorsRef = useRef(
    new Map<OfficeAnnotationTarget['kind'], OfficeAnnotationNavigator>(),
  );
  const onActivateRef = useRef(onActivate);
  snapshotRef.current = snapshot;
  onActivateRef.current = onActivate;

  const clearRuntimeState = useCallback(() => {
    readGenerationRef.current += 1;
    navigationGenerationRef.current += 1;
    revisionSelectionGenerationRef.current += 1;
    readControllerRef.current?.abort();
    readControllerRef.current = undefined;
    cacheRef.current.clear();
    setCacheRevision((current) => current + 1);
    setActiveIndex(-1);
    setActiveAnnotation(undefined);
    setActiveRevision(undefined);
    setLoading(false);
    setError(undefined);
  }, []);

  useEffect(() => {
    clearRuntimeState();
  }, [clearRuntimeState, sessionKey]);

  useEffect(() => {
    if (!enabled || !source) {
      setSnapshot(EMPTY_ANNOTATION_SNAPSHOT);
      clearRuntimeState();
      return undefined;
    }
    let lastRevision = -1;
    const syncSnapshot = () => {
      const next = source.getSnapshot();
      if (next.revision !== lastRevision) {
        lastRevision = next.revision;
        clearRuntimeState();
      }
      setSnapshot(next);
    };
    syncSnapshot();
    return source.subscribe(syncSnapshot);
  }, [clearRuntimeState, enabled, source]);

  useEffect(
    () => () => {
      readControllerRef.current?.abort();
    },
    [],
  );

  const registerSource = useCallback((nextSource: OfficeAnnotationSource) => {
    sourceRef.current = nextSource;
    setSource(nextSource);
    return () => {
      if (sourceRef.current !== nextSource) return;
      sourceRef.current = undefined;
      setSource(undefined);
    };
  }, []);

  const registerNavigator = useCallback(
    (
      kind: OfficeAnnotationTarget['kind'],
      navigator: OfficeAnnotationNavigator,
    ) => {
      navigatorsRef.current.set(kind, navigator);
      return () => {
        if (navigatorsRef.current.get(kind) === navigator) {
          navigatorsRef.current.delete(kind);
        }
      };
    },
    [],
  );

  const loadRange = useCallback(async (start: number, end: number) => {
    const activeSource = sourceRef.current;
    const activeSnapshot = snapshotRef.current;
    if (!activeSource || !activeSnapshot.count) return;
    const rangeStart = Math.min(
      activeSnapshot.count,
      Math.max(0, Math.trunc(start)),
    );
    const rangeEnd = Math.min(
      activeSnapshot.count,
      Math.max(rangeStart, Math.trunc(end)),
    );
    if (rangeStart >= rangeEnd) return;

    const allCached = Array.from(
      { length: rangeEnd - rangeStart },
      (_, offset) => rangeStart + offset,
    ).every((index) => cacheRef.current.has(index));
    if (allCached) return;

    const generation = ++readGenerationRef.current;
    readControllerRef.current?.abort();
    const controller = new AbortController();
    readControllerRef.current = controller;
    setLoading(true);
    setError(undefined);
    try {
      const items = await activeSource.getRange(
        rangeStart,
        rangeEnd,
        controller.signal,
      );
      if (
        generation !== readGenerationRef.current ||
        sourceRef.current !== activeSource ||
        activeSource.getSnapshot().revision !== activeSnapshot.revision
      ) {
        return;
      }
      items.forEach((annotation, offset) => {
        cacheRef.current.set(rangeStart + offset, annotation);
      });
      setCacheRevision((current) => current + 1);
    } catch (rangeError) {
      if (!isAbortError(rangeError)) {
        setError(
          rangeError instanceof Error ? rangeError.message : String(rangeError),
        );
      }
    } finally {
      if (generation === readGenerationRef.current) {
        setLoading(false);
        if (readControllerRef.current === controller) {
          readControllerRef.current = undefined;
        }
      }
    }
  }, []);

  const getCached = useCallback(
    (index: number) => cacheRef.current.get(index),
    [],
  );

  const loadAnnotationsByIds = useCallback(
    async (ids: readonly string[]) => {
      const activeSource = sourceRef.current;
      if (!activeSource || !ids.length) return [];
      const sourceRevision = activeSource.getSnapshot().revision;
      const uniqueIds = [...new Set(ids)];
      const indexes = await Promise.all(
        uniqueIds.map((id) => activeSource.findIndexById(id)),
      );
      if (
        sourceRef.current !== activeSource ||
        activeSource.getSnapshot().revision !== sourceRevision
      ) {
        return [];
      }
      const validIndexes = indexes.filter((index) => index >= 0);
      if (!validIndexes.length) return [];
      // 同一可见 Word 页面中的批注通常连续，合并范围读取可避免逐条往返。
      await loadRange(Math.min(...validIndexes), Math.max(...validIndexes) + 1);
      if (
        sourceRef.current !== activeSource ||
        activeSource.getSnapshot().revision !== sourceRevision
      ) {
        return [];
      }
      const annotationById = new Map(
        validIndexes.flatMap((index) => {
          const annotation = cacheRef.current.get(index);
          return annotation ? [[annotation.id, annotation] as const] : [];
        }),
      );
      return uniqueIds.flatMap((id) => {
        const annotation = annotationById.get(id);
        return annotation ? [annotation] : [];
      });
    },
    [loadRange],
  );

  const selectIndex = useCallback(
    async (requestedIndex: number) => {
      const count = snapshotRef.current.count;
      const index = normalizeAnnotationIndex(requestedIndex, count);
      if (index < 0) return false;
      await loadRange(index, index + 1);
      const annotation = cacheRef.current.get(index);
      if (!annotation) return false;
      const navigator = navigatorsRef.current.get(annotation.target.kind);
      const generation = ++navigationGenerationRef.current;
      const navigated = navigator ? await navigator(annotation) : true;
      if (!navigated || generation !== navigationGenerationRef.current) {
        return false;
      }
      setActiveIndex(index);
      setActiveAnnotation(annotation);
      setActiveRevision(undefined);
      onActivateRef.current?.();
      return true;
    },
    [loadRange],
  );

  const selectId = useCallback(
    async (id: string) => {
      const activeSource = sourceRef.current;
      if (!activeSource) return false;
      const index = await activeSource.findIndexById(id);
      return index < 0 ? false : selectIndex(index);
    },
    [selectIndex],
  );

  const previous = useCallback(() => {
    const count = snapshotRef.current.count;
    if (!count) return Promise.resolve(false);
    return selectIndex(activeIndex <= 0 ? count - 1 : activeIndex - 1);
  }, [activeIndex, selectIndex]);

  const next = useCallback(() => {
    const count = snapshotRef.current.count;
    if (!count) return Promise.resolve(false);
    return selectIndex(activeIndex < 0 ? 0 : (activeIndex + 1) % count);
  }, [activeIndex, selectIndex]);

  const activateRevisionId = useCallback(async (id: string) => {
    const activeSource = sourceRef.current;
    if (
      !activeSource?.findRevisionIndexById ||
      !activeSource.getRevisionRange
    ) {
      return false;
    }
    const sourceRevision = activeSource.getSnapshot().revision;
    const generation = ++revisionSelectionGenerationRef.current;
    const index = await activeSource.findRevisionIndexById(id);
    if (index < 0) return false;
    const [revision] = await activeSource.getRevisionRange(index, index + 1);
    if (
      !revision ||
      generation !== revisionSelectionGenerationRef.current ||
      sourceRef.current !== activeSource ||
      activeSource.getSnapshot().revision !== sourceRevision
    ) {
      return false;
    }
    setActiveRevision(revision);
    setActiveIndex(-1);
    setActiveAnnotation(undefined);
    return true;
  }, []);

  const clearActiveRevision = useCallback(() => {
    revisionSelectionGenerationRef.current += 1;
    setActiveRevision(undefined);
  }, []);

  const actions = useMemo<OfficeAnnotationControllerActions>(
    () => ({
      registerSource,
      registerNavigator,
      loadRange,
      getCached,
      loadAnnotationsByIds,
      selectIndex,
      selectId,
      previous,
      next,
      activateRevisionId,
      clearActiveRevision,
    }),
    [
      getCached,
      loadAnnotationsByIds,
      loadRange,
      next,
      previous,
      registerNavigator,
      registerSource,
      selectId,
      selectIndex,
      activateRevisionId,
      clearActiveRevision,
    ],
  );

  return {
    options: { showComments, showNotes },
    state: {
      snapshot,
      activeIndex,
      activeAnnotation,
      activeRevision,
      loading,
      error,
      cacheRevision,
    },
    actions,
  };
}
