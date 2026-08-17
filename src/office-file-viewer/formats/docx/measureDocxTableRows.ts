import type { DocxMeasuredBlock } from '../../services/docx/docxPagination';

type VerticalInterval = {
  bottom: number;
  top: number;
};

/** 测量表格行高与安全拆分位置，避免跨页时从文字或图片中间裁开。 */
export function measureDocxTableRows(
  element: HTMLElement,
  block: DocxMeasuredBlock['block'],
): Pick<DocxMeasuredBlock, 'rowBreakOffsets' | 'rowHeights'> {
  if (block.type !== 'table') return {};
  const table = Array.from(element.children).find(
    (child): child is HTMLTableElement => child.tagName === 'TABLE',
  );
  const body = table?.tBodies[0];
  const rows = Array.from(body?.rows ?? []);
  if (rows.length !== block.rows.length) return {};

  const rowHeights = rows.map((row) => row.getBoundingClientRect().height);
  const rowBreakOffsets = rows.map((row) => {
    const rowRect = row.getBoundingClientRect();
    const intervals: VerticalInterval[] = [];
    const elements = Array.from(
      row.querySelectorAll<HTMLElement>('span, img, svg, canvas'),
    ).filter(
      (candidate) =>
        candidate.tagName !== 'SPAN' || !candidate.querySelector('span'),
    );
    elements.forEach((candidate) => {
      Array.from(candidate.getClientRects()).forEach((rect) => {
        const top = Math.max(0, rect.top - rowRect.top);
        const bottom = Math.min(rowRect.height, rect.bottom - rowRect.top);
        if (bottom - top > 0.5) intervals.push({ bottom, top });
      });
    });
    intervals.sort((left, right) => left.top - right.top);
    const bands = intervals.reduce<VerticalInterval[]>((result, interval) => {
      const previous = result[result.length - 1];
      if (!previous || interval.top > previous.bottom + 1) {
        result.push({ ...interval });
      } else {
        previous.bottom = Math.max(previous.bottom, interval.bottom);
      }
      return result;
    }, []);
    return bands.slice(0, -1).flatMap((band, index) => {
      const next = bands[index + 1];
      const offset = (band.bottom + next.top) / 2;
      return offset > 4 && offset < rowRect.height - 4 ? [offset] : [];
    });
  });
  return { rowBreakOffsets, rowHeights };
}
