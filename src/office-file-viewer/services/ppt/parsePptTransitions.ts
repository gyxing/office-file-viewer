import {
  normalizePresentationTransitionDuration,
  type PresentationTransition,
  type PresentationTransitionDirection,
} from '../presentation/transitionTypes';
import type { PresentationWarning } from '../presentation/types';
import { PPT_RECORD } from './binary/constants';
import { walkPptRecords } from './binary/walkPptRecords';
import type { PptRecord } from './types';

const DIRECTION: Record<number, PresentationTransitionDirection> = {
  0: 'left',
  1: 'up',
  2: 'right',
  3: 'down',
};

/** 从 SlideShowSlideInfoAtom 恢复当前支持的二进制 PPT 页级切换。 */
export function parsePptTransition(
  slideRecord: PptRecord,
  slideIndex: number,
): { transition?: PresentationTransition; warning?: PresentationWarning } {
  let transition: PresentationTransition | undefined;
  let warning: PresentationWarning | undefined;
  walkPptRecords(slideRecord, (record) => {
    if (
      transition ||
      warning ||
      record.type !== PPT_RECORD.SLIDE_SHOW_SLIDE_INFO_ATOM ||
      record.data.length < 13
    ) {
      return;
    }
    const direction = record.data[8];
    const effect = record.data[9];
    const speed = record.data[12];
    const durationMs = normalizePresentationTransitionDuration(
      speed === 0 ? 750 : speed === 2 ? 250 : 500,
    );
    if (effect === 6 || effect === 23) {
      transition = { type: 'fade', durationMs };
      return;
    }
    if ([4, 7, 10, 20].includes(effect)) {
      const type =
        effect === 4
          ? 'cover'
          : effect === 7
          ? 'uncover'
          : effect === 10
          ? 'wipe'
          : 'push';
      transition = {
        type,
        direction: DIRECTION[direction] ?? 'left',
        durationMs,
      };
      return;
    }
    if (effect === 13) {
      const splitDirections: PresentationTransitionDirection[] = [
        'horizontal-out',
        'horizontal-in',
        'vertical-out',
        'vertical-in',
      ];
      transition = {
        type: 'split',
        direction: splitDirections[direction] ?? 'horizontal-out',
        durationMs,
      };
      return;
    }
    if (effect !== 0 && effect !== 255) {
      warning = {
        code: 'PPT_TRANSITION_UNSUPPORTED',
        message: `暂不支持编号 ${effect} 的 PPT 幻灯片切换，已使用静态切页`,
        slideIndex,
      };
    }
  });
  return { transition, warning };
}
