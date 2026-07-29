import type { WordPageMeta } from '../../services/word/WordPageSource';

export type WordBlockPageLocation = {
  pageIndex: number;
  pageRevision: number;
};

/** 维护正文块到最新页面 revision 的映射。 */
export class WordBlockPageIndex {
  private readonly locations = new Map<string, WordBlockPageLocation>();
  private readonly pageEntries = new Map<
    number,
    { revision: number; blockIds: readonly string[] }
  >();
  private currentRevision = 0;

  get revision() {
    return this.currentRevision;
  }

  locate(blockId: string) {
    return this.locations.get(blockId);
  }

  replacePage(meta: WordPageMeta) {
    const previous = this.pageEntries.get(meta.index);
    if (previous && previous.revision > meta.revision) return;
    if (previous) this.removePage(meta.index, previous.revision);
    const blockIds = [...new Set(meta.sourceBlockIds)];
    blockIds.forEach((blockId) => {
      this.locations.set(blockId, {
        pageIndex: meta.index,
        pageRevision: meta.revision,
      });
    });
    this.pageEntries.set(meta.index, {
      revision: meta.revision,
      blockIds,
    });
    this.currentRevision += 1;
  }

  removePage(pageIndex: number, pageRevision: number) {
    const current = this.pageEntries.get(pageIndex);
    if (!current || current.revision !== pageRevision) return;
    current.blockIds.forEach((blockId) => {
      const location = this.locations.get(blockId);
      if (
        location?.pageIndex === pageIndex &&
        location.pageRevision === pageRevision
      ) {
        this.locations.delete(blockId);
      }
    });
    this.pageEntries.delete(pageIndex);
    this.currentRevision += 1;
  }
}
