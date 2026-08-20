import { attr } from '../../shared/ooxml/xml';
import {
  normalizePresentationTransitionDuration,
  type PresentationTransition,
  type PresentationTransitionDirection,
} from '../presentation/transitionTypes';
import type { PresentationWarning } from '../presentation/types';

/** 单页 PPTX 切换解析结果。 */
export type PptxTransitionParseResult = Readonly<{
  /** 当前查看器能够还原的切换。 */
  transition?: PresentationTransition;
  /** 源切换不在支持子集时的降级说明。 */
  warning?: PresentationWarning;
}>;

function normalizeDirection(value?: string): PresentationTransitionDirection {
  if (value === 'r') return 'right';
  if (value === 'u') return 'up';
  if (value === 'd') return 'down';
  return 'left';
}

function readDuration(node: Element) {
  const explicit = Number(attr(node, 'dur'));
  if (Number.isFinite(explicit) && explicit > 0) return explicit;
  const speed = attr(node, 'spd');
  return speed === 'slow' ? 1000 : speed === 'fast' ? 250 : 500;
}

/** 解析 fade、push、wipe、split、cover 和 uncover 页级切换。 */
export function parsePptxTransition(
  transitionNode: Element | null,
  slideIndex: number,
): PptxTransitionParseResult {
  if (!transitionNode) return {};
  const effect = Array.from(transitionNode.children)[0];
  if (!effect) return {};
  const name = effect.localName.toLowerCase();
  const durationMs = normalizePresentationTransitionDuration(
    readDuration(transitionNode),
  );
  if (name === 'fade') return { transition: { type: 'fade', durationMs } };
  if (['push', 'wipe', 'cover', 'uncover'].includes(name)) {
    return {
      transition: {
        type: name as 'push' | 'wipe' | 'cover' | 'uncover',
        direction: normalizeDirection(attr(effect, 'dir')),
        durationMs,
      },
    };
  }
  if (name === 'split') {
    const horizontal = attr(effect, 'orient') !== 'vert';
    const inward = attr(effect, 'dir') === 'in';
    return {
      transition: {
        type: 'split',
        direction: `${horizontal ? 'horizontal' : 'vertical'}-${
          inward ? 'in' : 'out'
        }`,
        durationMs,
      },
    };
  }
  return {
    warning: {
      code: 'PPTX_TRANSITION_UNSUPPORTED',
      message: `暂不支持 ${effect.localName} 幻灯片切换，已使用静态切页`,
      slideIndex,
    },
  };
}
