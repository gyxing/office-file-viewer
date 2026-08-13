import { useLayoutEffect, useRef, useState } from 'react';
import type { TextStyle } from '../../../services/pptx/types';

/** 浏览器字体度量导致内容超出时，每轮使用的渐进缩放比例。 */
const AUTO_FIT_STEP = 0.95;

/** 防止异常文档把文字无限缩小。 */
const MIN_AUTO_FIT_SCALE = 0.5;

/** 补偿 Office 与 Chromium 在常见中文字体上的字符推进宽度差异。 */
const AUTO_FIT_TRACKING_COMPENSATION = -1;

/** 计算普通流段落实际占用的高度，排除绝对定位的倒影副本。 */
function measureParagraphHeight(paragraphs: HTMLElement[]) {
  return paragraphs.reduce((height, paragraph) => {
    const style = getComputedStyle(paragraph);
    return (
      height +
      paragraph.offsetHeight +
      (Number.parseFloat(style.marginTop) || 0) +
      (Number.parseFloat(style.marginBottom) || 0)
    );
  }, 0);
}

/** 在主视口中渐进缩小需要补偿的自动适应文本，缩略图不做批量 DOM 测量。 */
export function usePresentationTextAutoFit(
  fit: TextStyle['fit'],
  enabled: boolean,
  identity: string,
) {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const [scale, setScale] = useState(1);
  const [tracking, setTracking] = useState(0);

  useLayoutEffect(() => {
    setScale(1);
    setTracking(0);
  }, [identity]);

  useLayoutEffect(() => {
    const container = containerRef.current;
    if (!container || !enabled || !fit || fit === 'none') return;

    const paragraphs = Array.from(
      container.querySelectorAll<HTMLElement>(
        ':scope > [data-office-presentation-text-paragraph="true"]',
      ),
    );
    if (!paragraphs.length) return;

    const containerStyle = getComputedStyle(container);
    const availableHeight =
      container.clientHeight -
      (Number.parseFloat(containerStyle.paddingTop) || 0) -
      (Number.parseFloat(containerStyle.paddingBottom) || 0);
    // 扩大形状模式允许浏览器字体取整产生少量误差，避免把源文件正确字号误缩小。
    const overflowTolerance = fit === 'resizeShape' ? 4 : 1;
    const heightOverflow =
      measureParagraphHeight(paragraphs) >
      availableHeight + overflowTolerance;
    const widthOverflow = paragraphs.some(
      (paragraph) =>
        paragraph.scrollWidth > paragraph.clientWidth + overflowTolerance,
    );
    if (!heightOverflow && !widthOverflow) return;

    if (fit === 'resizeShape' && tracking === 0) {
      setTracking(AUTO_FIT_TRACKING_COMPENSATION);
      return;
    }

    setScale((current) =>
      current <= MIN_AUTO_FIT_SCALE
        ? current
        : Math.max(
            MIN_AUTO_FIT_SCALE,
            Math.round(current * AUTO_FIT_STEP * 1000) / 1000,
          ),
    );
  }, [enabled, fit, identity, scale, tracking]);

  return { containerRef, scale, tracking };
}
