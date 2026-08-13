import React, { memo, useEffect, useMemo, useRef, useState } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type { OfficeSearchResult } from '../../services/search/types';

/** 结果较多时才启用固定行高窗口渲染，普通文档保留完整 DOM。 */
const SEARCH_RESULT_VIRTUAL_THRESHOLD = 400;
/** 单条结果固定高度，避免虚拟列表滚动位置随内容抖动。 */
const SEARCH_RESULT_ROW_HEIGHT = 62;
/** 可视区上下额外保留的结果行数。 */
const SEARCH_RESULT_OVERSCAN = 6;

/** 搜索结果列表属性。 */
type OfficeSearchResultListProps = {
  /** 当前查询的全部结果。 */
  results: readonly OfficeSearchResult[];
  /** 当前完成导航的结果索引。 */
  currentIndex: number;
  /** 请求定位指定结果。 */
  onSelect(index: number): void;
};

function spreadsheetColumnLabel(columnIndex: number) {
  let current = Math.max(1, Math.trunc(columnIndex));
  let label = '';
  while (current > 0) {
    current -= 1;
    label = String.fromCharCode(65 + (current % 26)) + label;
    current = Math.floor(current / 26);
  }
  return label;
}

/** 渲染单条搜索结果，并保留格式相关的位置摘要。 */
function SearchResultRow({
  result,
  index,
  selected,
  onSelect,
}: {
  /** 当前匹配结果。 */
  result: OfficeSearchResult;
  /** 当前结果在完整集合中的索引。 */
  index: number;
  /** 当前结果是否已经完成导航。 */
  selected: boolean;
  /** 请求定位当前结果。 */
  onSelect(index: number): void;
}) {
  const messages = useOfficeFileViewerMessages();
  let location: string;
  if (result.target.kind === 'spreadsheet') {
    location = `${spreadsheetColumnLabel(result.target.columnIndex)}${
      result.target.rowIndex
    }`;
  } else if (result.target.kind === 'presentation') {
    location = messages.presentation.slide(result.target.slideIndex + 1);
  } else {
    location =
      result.target.pageIndex === undefined
        ? messages.search.title
        : messages.search.page(result.target.pageIndex + 1);
  }

  return (
    <button
      type="button"
      className="office-file-search-results__item"
      role="option"
      aria-selected={selected}
      data-result-index={index}
      data-testid="office-search-result"
      onClick={() => onSelect(index)}
    >
      <span className="office-file-search-results__meta">
        <span>{location}</span>
        {result.target.kind === 'presentation' && result.target.hidden ? (
          <span className="office-file-search-results__hidden">
            {messages.search.hiddenSlide}
          </span>
        ) : null}
      </span>
      <span
        className="office-file-search-results__preview"
        title={result.previewText}
      >
        {result.previewText}
      </span>
    </button>
  );
}

/** 在大结果集时按窗口渲染，并让当前项尽量位于列表中部。 */
function OfficeSearchResultListComponent({
  results,
  currentIndex,
  onSelect,
}: OfficeSearchResultListProps) {
  const viewportRef = useRef<HTMLDivElement>(null);
  const [scrollTop, setScrollTop] = useState(0);
  const [viewportHeight, setViewportHeight] = useState(0);
  const virtual = results.length >= SEARCH_RESULT_VIRTUAL_THRESHOLD;

  useEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return;
    const update = () => setViewportHeight(viewport.clientHeight);
    update();
    const observer =
      typeof ResizeObserver === 'undefined'
        ? undefined
        : new ResizeObserver(update);
    observer?.observe(viewport);
    return () => observer?.disconnect();
  }, []);

  useEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport || currentIndex < 0 || currentIndex >= results.length) return;
    const targetTop = currentIndex * SEARCH_RESULT_ROW_HEIGHT;
    const centeredTop =
      targetTop -
      Math.max(0, (viewport.clientHeight - SEARCH_RESULT_ROW_HEIGHT) / 2);
    const maximumTop = Math.max(
      0,
      results.length * SEARCH_RESULT_ROW_HEIGHT - viewport.clientHeight,
    );
    const nextTop = Math.min(maximumTop, Math.max(0, centeredTop));
    viewport.scrollTo({
      top: nextTop,
      behavior: window.matchMedia?.('(prefers-reduced-motion: reduce)').matches
        ? 'auto'
        : 'smooth',
    });
  }, [currentIndex, results.length]);

  const windowRange = useMemo(() => {
    if (!virtual) return { start: 0, end: results.length };
    const start = Math.max(
      0,
      Math.floor(scrollTop / SEARCH_RESULT_ROW_HEIGHT) - SEARCH_RESULT_OVERSCAN,
    );
    const visibleCount = Math.ceil(
      Math.max(SEARCH_RESULT_ROW_HEIGHT, viewportHeight) /
        SEARCH_RESULT_ROW_HEIGHT,
    );
    return {
      start,
      end: Math.min(
        results.length,
        start + visibleCount + SEARCH_RESULT_OVERSCAN * 2,
      ),
    };
  }, [results.length, scrollTop, viewportHeight, virtual]);
  const visibleResults = results.slice(windowRange.start, windowRange.end);

  return (
    <div
      ref={viewportRef}
      className="office-file-search-results"
      role="listbox"
      onScroll={(event) => setScrollTop(event.currentTarget.scrollTop)}
    >
      <div
        className="office-file-search-results__canvas"
        data-virtual={virtual ? 'true' : 'false'}
        style={
          virtual
            ? { height: results.length * SEARCH_RESULT_ROW_HEIGHT }
            : undefined
        }
      >
        <div
          className="office-file-search-results__window"
          style={
            virtual
              ? {
                  transform: `translateY(${
                    windowRange.start * SEARCH_RESULT_ROW_HEIGHT
                  }px)`,
                }
              : undefined
          }
        >
          {visibleResults.map((result, offset) => {
            const index = windowRange.start + offset;
            return (
              <SearchResultRow
                key={result.id}
                result={result}
                index={index}
                selected={index === currentIndex}
                onSelect={onSelect}
              />
            );
          })}
        </div>
      </div>
    </div>
  );
}

export const OfficeSearchResultList = memo(OfficeSearchResultListComponent);
