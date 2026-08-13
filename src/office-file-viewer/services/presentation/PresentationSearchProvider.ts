import {
  OfficeSearchBatchWriter,
  throwIfOfficeSearchAborted,
} from '../search/OfficeSearchProvider';
import {
  createSearchPreviewText,
  findSearchMatches,
  normalizeSearchText,
} from '../search/normalizeSearchText';
import type {
  OfficeSearchProgressEmitter,
  OfficeSearchProvider,
  OfficeSearchQuery,
  OfficeSearchResult,
} from '../search/types';
import type { PresentationSource } from './PresentationSource';
import type { SlideElement, TextParagraph } from './types';

/** 幻灯片搜索使用的可见元素文本。 */
type PresentationSearchText = Readonly<{
  elementId: string;
  text: string;
}>;

function paragraphText(paragraphs: readonly TextParagraph[]) {
  return paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n');
}

function collectElementText(element: SlideElement): PresentationSearchText[] {
  if (element.opacity === 0) return [];
  if (element.type === 'text') {
    return [{ elementId: element.id, text: paragraphText(element.paragraphs) }];
  }
  if (element.type === 'table') {
    return [
      {
        elementId: element.id,
        text: element.rows
          .map((row) => row.map((cell) => cell.text).join('\t'))
          .join('\n'),
      },
    ];
  }
  if (element.type === 'group') {
    return element.children.flatMap(collectElementText);
  }
  return [];
}

/** 按幻灯片和元素绘制顺序扫描正文，不读取备注或批注。 */
export class PresentationSearchProvider implements OfficeSearchProvider {
  readonly kind = 'presentation' as const;

  constructor(private readonly source: PresentationSource) {}

  async search(
    query: OfficeSearchQuery,
    emit: OfficeSearchProgressEmitter,
    signal: AbortSignal,
  ) {
    throwIfOfficeSearchAborted(signal);
    const snapshot = this.source.getSnapshot();
    const estimatedTotal = snapshot.slides.reduce(
      (total, slide) => total + Math.max(0, slide.estimatedElementCount),
      0,
    );
    const writer = new OfficeSearchBatchWriter(emit, signal, estimatedTotal);
    if (!normalizeSearchText(query.text, query.matchCase).text) {
      writer.complete();
      return;
    }

    for (
      let slideIndex = 0;
      slideIndex < snapshot.slideCount;
      slideIndex += 1
    ) {
      throwIfOfficeSearchAborted(signal);
      const descriptor = snapshot.slides[slideIndex];
      const slide = await this.source.getSlide(slideIndex, signal);
      const searchableElements = slide.elements.flatMap(collectElementText);
      writer.setTotal(Math.max(estimatedTotal, searchableElements.length));

      for (const element of searchableElements) {
        const items: OfficeSearchResult[] = findSearchMatches(
          element.text,
          query,
        ).map(({ startOffset, endOffset }) => ({
          id: `presentation:${slideIndex}:${element.elementId}:${startOffset}:${endOffset}`,
          matchText: element.text.slice(startOffset, endOffset),
          previewText: createSearchPreviewText(
            element.text,
            startOffset,
            endOffset,
          ),
          target: {
            kind: 'presentation',
            slideIndex,
            elementId: element.elementId,
            startOffset,
            endOffset,
            hidden: Boolean(descriptor?.hidden ?? slide.hidden),
          },
        }));
        await writer.append(items);
      }
    }
    writer.complete();
  }
}
