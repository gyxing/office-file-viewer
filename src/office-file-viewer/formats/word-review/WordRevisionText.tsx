import type { MouseEvent, ReactNode } from 'react';
import React, { createContext, useContext } from 'react';
import type { WordRevisionMode } from '../../services/annotations/types';
import type { WordInlineReview } from '../../services/word/review/types';
import { useOfficeAnnotationRuntime } from '../../shared/annotations';
import { getWordRevisionCssVariables } from './wordRevisionAppearance';

/** 当前 DOCX 树采用的修订投影模式。 */
export const WordRevisionModeContext = createContext<WordRevisionMode>('final');

/** 读取当前 DOCX 修订投影模式。 */
export function useWordRevisionMode() {
  return useContext(WordRevisionModeContext);
}

/** 标记态行内修订包装属性。 */
type WordRevisionTextProps = {
  /** 当前文字关联的一个或多个修订。 */
  review?: WordInlineReview;
  /** 已由格式渲染器生成的行内内容。 */
  children: ReactNode;
};

/** 在 markup 模式中标记插入、删除、移动和格式变化。 */
export function WordRevisionText({ review, children }: WordRevisionTextProps) {
  const mode = useWordRevisionMode();
  const runtime = useOfficeAnnotationRuntime();
  const revisions = review?.revisions ?? [];
  if (mode !== 'markup' || !revisions.length) return <>{children}</>;
  const kinds = new Set(revisions.map((revision) => revision.kind));
  const className = [
    'office-file-word-revision',
    kinds.has('insert') || kinds.has('move-to')
      ? 'office-file-word-revision--insert'
      : undefined,
    kinds.has('delete') || kinds.has('move-from')
      ? 'office-file-word-revision--delete'
      : undefined,
    kinds.has('format') ? 'office-file-word-revision--format' : undefined,
  ]
    .filter(Boolean)
    .join(' ');
  const latestRevision = revisions[revisions.length - 1];
  const active = runtime?.state.activeRevision?.id === latestRevision.id;
  const handleClick = (event: MouseEvent<HTMLSpanElement>) => {
    // Ctrl/Command + 单击优先保留超链接跳转，不用修订选择覆盖用户意图。
    if (event.ctrlKey || event.metaKey || event.defaultPrevented) return;
    event.stopPropagation();
    if (active) runtime?.actions.clearActiveRevision();
    else void runtime?.actions.activateRevisionId(latestRevision.id);
  };
  return (
    <span
      className={className}
      data-office-word-revision={latestRevision.id}
      data-office-word-revision-kind={latestRevision.kind}
      data-office-word-revision-author={latestRevision.author}
      data-office-word-revision-created-at={latestRevision.createdAt}
      data-active={active ? 'true' : 'false'}
      style={getWordRevisionCssVariables(latestRevision)}
      title={latestRevision.author}
      onClick={handleClick}
    >
      {children}
    </span>
  );
}
