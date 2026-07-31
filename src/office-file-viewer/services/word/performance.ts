import { OFFICE_LARGE_FILE_THRESHOLDS } from '../performance/officePerformanceThresholds';
import type { WordPerformanceProfile, WordPerformanceStats } from './types';

/** 启用 Word 大纲虚拟列表的节点数量阈值。 */
export const WORD_OUTLINE_VIRTUAL_THRESHOLD = 200;
/** 启用 Word 页面虚拟窗口的页数阈值。 */
export const WORD_PAGE_WINDOW_THRESHOLD = 60;
/** 启用 Word 按需渲染的内容权重阈值。 */
export const WORD_RENDER_WEIGHT_THRESHOLD = 3000;
/** 折算一个 Word 文本渲染权重所需的字符数。 */
export const WORD_TEXT_WEIGHT_CHARACTERS = 2000;

/** 按统一权重计算 Word 文档渲染成本。 */
export function calculateWordRenderWeight(stats: WordPerformanceStats) {
  return (
    stats.paragraphCount +
    stats.tableRowCount * 4 +
    stats.imageCount * 20 +
    stats.drawingCount * 30 +
    Math.ceil(stats.textLength / WORD_TEXT_WEIGHT_CHARACTERS)
  );
}

/** 由轻量统计生成普通或大文件单向升级所需的初始画像。 */
export function createWordPerformanceProfile(
  stats: WordPerformanceStats,
): WordPerformanceProfile {
  const renderWeight = calculateWordRenderWeight(stats);
  return {
    renderWeight,
    outlineMode:
      stats.outlineCount > WORD_OUTLINE_VIRTUAL_THRESHOLD
        ? 'virtual'
        : 'normal',
    pageMode:
      (stats.estimatedPageCount ?? 0) >= WORD_PAGE_WINDOW_THRESHOLD ||
      renderWeight >= WORD_RENDER_WEIGHT_THRESHOLD ||
      stats.slowPagination
        ? 'windowed'
        : 'normal',
  };
}

/** 判断一次分页或测量批次是否需要触发页面窗口升级。 */
export function isSlowWordPagination(durationMs: number) {
  return durationMs >= OFFICE_LARGE_FILE_THRESHOLDS.slowTaskMilliseconds;
}
