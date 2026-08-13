import type { ReactNode } from 'react';
import React, {
  createContext,
  memo,
  useContext,
  useEffect,
  useMemo,
} from 'react';
import { findSearchMatches } from '../../services/search/normalizeSearchText';
import type {
  OfficeSearchProvider,
  OfficeSearchTarget,
} from '../../services/search/types';
import type {
  OfficeSearchController,
  OfficeSearchNavigator,
} from './useOfficeSearchController';

/** 文本渲染器定位可见匹配时使用的轻量目标。 */
export type OfficeSearchHighlightTarget =
  | Readonly<{ kind: 'word'; blockId: string }>
  | Readonly<{
      kind: 'spreadsheet';
      sheetId: string;
      rowIndex: number;
      columnIndex: number;
    }>
  | Readonly<{
      kind: 'presentation';
      slideIndex: number;
      elementId: string;
    }>;

type OfficeSearchRuntimeValue = Readonly<{
  controller: OfficeSearchController;
  matchingTargetKeys: ReadonlySet<string>;
  currentTargetKey?: string;
}>;

const OfficeSearchRuntimeContext = createContext<
  OfficeSearchRuntimeValue | undefined
>(undefined);

/** 返回忽略字符偏移后的渲染目标键，用于只处理当前可见节点。 */
export function getOfficeSearchTargetKey(
  target: OfficeSearchTarget | OfficeSearchHighlightTarget,
) {
  switch (target.kind) {
    case 'word':
      return `word:${target.blockId}`;
    case 'spreadsheet':
      return `spreadsheet:${target.sheetId}:${target.rowIndex}:${target.columnIndex}`;
    case 'presentation':
      return `presentation:${target.slideIndex}:${target.elementId}`;
  }
}

/** 为搜索侧栏和各格式渲染器提供同一控制器实例。 */
function OfficeSearchRuntimeProviderComponent({
  controller,
  children,
}: {
  /** 当前查看器独占的搜索控制器。 */
  controller: OfficeSearchController;
  /** 工具栏、侧栏与格式查看器。 */
  children: ReactNode;
}) {
  const matchingTargetKeys = useMemo(
    () =>
      new Set(
        controller.state.results.map((result) =>
          getOfficeSearchTargetKey(result.target),
        ),
      ),
    [controller.state.results],
  );
  const currentTarget =
    controller.state.currentIndex >= 0
      ? controller.state.results[controller.state.currentIndex]?.target
      : undefined;
  const value = useMemo<OfficeSearchRuntimeValue>(
    () => ({
      controller,
      matchingTargetKeys,
      currentTargetKey: currentTarget
        ? getOfficeSearchTargetKey(currentTarget)
        : undefined,
    }),
    [controller, currentTarget, matchingTargetKeys],
  );
  return (
    <OfficeSearchRuntimeContext.Provider value={value}>
      {children}
    </OfficeSearchRuntimeContext.Provider>
  );
}

export const OfficeSearchRuntimeProvider = memo(
  OfficeSearchRuntimeProviderComponent,
);

/** 仅在搜索启用时建立运行时作用域，关闭能力后不让格式查看器注册提供器。 */
export function OfficeSearchRuntimeBoundary({
  enabled,
  controller,
  children,
}: {
  /** 当前查看器是否启用搜索能力。 */
  enabled: boolean;
  /** 当前查看器独占的搜索控制器。 */
  controller: OfficeSearchController;
  /** 工具栏、侧栏与格式查看器。 */
  children: ReactNode;
}) {
  if (!enabled) return <>{children}</>;
  return (
    <OfficeSearchRuntimeProvider controller={controller}>
      {children}
    </OfficeSearchRuntimeProvider>
  );
}

/** 读取当前查看器的搜索运行时；未启用搜索时返回空值。 */
export function useOfficeSearchRuntime() {
  return useContext(OfficeSearchRuntimeContext);
}

/** 在格式查看器挂载期间注册其搜索数据提供器。 */
export function useOfficeSearchProviderRegistration(
  provider: OfficeSearchProvider | undefined,
) {
  const runtime = useOfficeSearchRuntime();
  const registerProvider = runtime?.controller.actions.registerProvider;
  useEffect(() => {
    if (!provider || !registerProvider) return undefined;
    return registerProvider(provider);
  }, [provider, registerProvider]);
}

/** 在格式查看器挂载期间注册精确结果导航能力。 */
export function useOfficeSearchNavigatorRegistration(
  kind: OfficeSearchTarget['kind'],
  navigator: OfficeSearchNavigator | undefined,
) {
  const runtime = useOfficeSearchRuntime();
  const registerNavigator = runtime?.controller.actions.registerNavigator;
  useEffect(() => {
    if (!navigator || !registerNavigator) return undefined;
    return registerNavigator(kind, navigator);
  }, [kind, navigator, registerNavigator]);
}

/** 按当前查询拆分文本节点并渲染不改变排版尺寸的高亮标记。 */
export function OfficeSearchHighlightedText({
  text,
  target,
  renderText = (value) => value,
}: {
  /** 当前可见文本节点的原始内容。 */
  text: string;
  /** 当前文本对应的格式定位目标。 */
  target: OfficeSearchHighlightTarget;
  /** 保留压缩标点等格式专属文本渲染方式。 */
  renderText?: (value: string) => ReactNode;
}) {
  const runtime = useOfficeSearchRuntime();
  const targetKey = getOfficeSearchTargetKey(target);
  if (
    !runtime ||
    !runtime.matchingTargetKeys.has(targetKey) ||
    !runtime.controller.state.query
  ) {
    return <>{renderText(text)}</>;
  }
  const matches = findSearchMatches(text, {
    text: runtime.controller.state.query,
    matchCase: runtime.controller.state.matchCase,
    wholeWord: runtime.controller.state.wholeWord,
  });
  if (!matches.length) return <>{renderText(text)}</>;

  const current = runtime.currentTargetKey === targetKey;
  const nodes: ReactNode[] = [];
  let offset = 0;
  matches.forEach((match, index) => {
    if (match.startOffset > offset) {
      nodes.push(
        <React.Fragment key={`text-${index}`}>
          {renderText(text.slice(offset, match.startOffset))}
        </React.Fragment>,
      );
    }
    nodes.push(
      <mark
        key={`match-${index}`}
        className={
          current
            ? 'office-file-search-highlight office-file-search-highlight--current'
            : 'office-file-search-highlight'
        }
        data-office-search-highlight="true"
      >
        {renderText(text.slice(match.startOffset, match.endOffset))}
      </mark>,
    );
    offset = match.endOffset;
  });
  if (offset < text.length) {
    nodes.push(
      <React.Fragment key="text-tail">
        {renderText(text.slice(offset))}
      </React.Fragment>,
    );
  }
  return <>{nodes}</>;
}
