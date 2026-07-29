import { useCallback, useEffect, useMemo, useState } from 'react';
import {
  createWordPerformanceProfile,
  isSlowWordPagination,
} from '../../services/word/performance';
import type {
  WordPerformanceProfile,
  WordPerformanceStats,
} from '../../services/word/types';

type ProfileState = {
  documentSessionId: string;
  profile: WordPerformanceProfile;
};

function mergeProfile(
  current: WordPerformanceProfile,
  next: WordPerformanceProfile,
): WordPerformanceProfile {
  return {
    renderWeight: Math.max(current.renderWeight, next.renderWeight),
    outlineMode:
      current.outlineMode === 'virtual' || next.outlineMode === 'virtual'
        ? 'virtual'
        : 'normal',
    pageMode:
      current.pageMode === 'windowed' || next.pageMode === 'windowed'
        ? 'windowed'
        : 'normal',
  };
}

/** 根据 Session 隔离 Word 性能模式，并保证同一文档只允许单向升级。 */
export function useWordPerformanceProfile(
  documentSessionId: string,
  stats: WordPerformanceStats,
) {
  const nextProfile = useMemo(
    () => createWordPerformanceProfile(stats),
    [stats],
  );
  const [state, setState] = useState<ProfileState>({
    documentSessionId,
    profile: nextProfile,
  });
  const currentProfile =
    state.documentSessionId === documentSessionId
      ? mergeProfile(state.profile, nextProfile)
      : nextProfile;

  useEffect(() => {
    setState((current) => {
      const next =
        current.documentSessionId === documentSessionId
          ? mergeProfile(current.profile, nextProfile)
          : nextProfile;
      return current.documentSessionId === documentSessionId &&
        current.profile.renderWeight === next.renderWeight &&
        current.profile.outlineMode === next.outlineMode &&
        current.profile.pageMode === next.pageMode
        ? current
        : { documentSessionId, profile: next };
    });
  }, [documentSessionId, nextProfile]);

  const reportPaginationDuration = useCallback(
    (durationMs: number) => {
      if (!isSlowWordPagination(durationMs)) return;
      setState((current) => {
        const base =
          current.documentSessionId === documentSessionId
            ? current.profile
            : nextProfile;
        if (base.pageMode === 'windowed') return current;
        return {
          documentSessionId,
          profile: { ...base, pageMode: 'windowed' },
        };
      });
    },
    [documentSessionId, nextProfile],
  );

  return { profile: currentProfile, reportPaginationDuration };
}
