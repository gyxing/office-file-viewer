import type { DocBlock, DocDocument, DocTextInline } from '../doc/types';
import type {
  DocxBlock,
  DocxDocument,
  DocxInline,
  DocxShape,
} from '../docx/types';
import type { WordPerformanceStats } from './types';

type MutableStats = Omit<WordPerformanceStats, 'imageCount' | 'drawingCount'>;

/** 增量收集 Word 渲染权重，资源按稳定 ID 去重。 */
export class WordPerformanceStatsCollector {
  private readonly imageIds = new Set<string>();
  private readonly drawingIds = new Set<string>();
  private readonly visitedBlockIds = new Set<string>();
  private stats: MutableStats;

  constructor(initial: Partial<WordPerformanceStats> = {}) {
    this.stats = {
      estimatedPageCount: initial.estimatedPageCount,
      outlineCount: initial.outlineCount ?? 0,
      paragraphCount: initial.paragraphCount ?? 0,
      tableRowCount: initial.tableRowCount ?? 0,
      textLength: initial.textLength ?? 0,
      largestXmlSize: initial.largestXmlSize,
      slowPagination: initial.slowPagination ?? false,
    };
  }

  addImage(id: string) {
    if (id) this.imageIds.add(id);
  }

  addDrawing(id: string) {
    if (id) this.drawingIds.add(id);
  }

  addDocBlocks(blocks: readonly DocBlock[]) {
    blocks.forEach((block) => this.visitDocBlock(block));
  }

  addDocxBlocks(blocks: readonly DocxBlock[]) {
    blocks.forEach((block) => this.visitDocxBlock(block));
  }

  setEstimatedPageCount(count: number | undefined) {
    this.stats.estimatedPageCount = count;
  }

  setOutlineCount(count: number) {
    this.stats.outlineCount = count;
  }

  setLargestXmlSize(size: number | undefined) {
    this.stats.largestXmlSize = size;
  }

  reportSlowPagination() {
    this.stats.slowPagination = true;
  }

  getSnapshot(): WordPerformanceStats {
    return {
      ...this.stats,
      imageCount: this.imageIds.size,
      drawingCount: this.drawingIds.size,
    };
  }

  private addText(text: string | undefined) {
    this.stats.textLength += text?.length ?? 0;
  }

  private visitDocInlines(inlines: readonly DocTextInline[] | undefined) {
    inlines?.forEach((inline) => {
      if (inline.type === 'image') this.addImage(inline.image.id);
    });
  }

  private visitDocBlock(block: DocBlock) {
    if (this.visitedBlockIds.has(block.id)) return;
    this.visitedBlockIds.add(block.id);
    if (block.type === 'paragraph') {
      this.stats.paragraphCount += 1;
      this.addText(block.text);
      this.visitDocInlines(block.inlines);
      return;
    }
    if (block.type === 'list') {
      this.stats.paragraphCount += block.items.length;
      block.items.forEach((item) => {
        this.addText(item.text);
        this.visitDocInlines(item.inlines);
      });
      return;
    }
    this.stats.tableRowCount += block.rows.length;
    block.rows.forEach((row) =>
      row.cells.forEach((cell) => {
        this.addText(cell.text);
        this.visitDocInlines(cell.inlines);
      }),
    );
  }

  private visitDocxInlines(inlines: readonly DocxInline[]) {
    inlines.forEach((inline) => {
      if (inline.type === 'image') this.addImage(inline.image.id);
      else if (inline.type === 'chart') {
        this.addDrawing(inline.chart.id);
      } else if (inline.type === 'shape') {
        this.visitDocxShape(inline.shape);
      }
    });
  }

  private visitDocxShape(shape: DocxShape) {
    this.addDrawing(shape.id);
    shape.items.forEach((item) => {
      if (item.imageSrc) this.addImage(item.id);
      item.blocks?.forEach((block) => this.visitDocxBlock(block));
      item.paragraphs?.forEach((block) => this.visitDocxBlock(block));
    });
  }

  private visitDocxBlock(block: DocxBlock) {
    if (this.visitedBlockIds.has(block.id)) return;
    this.visitedBlockIds.add(block.id);
    if (block.type === 'paragraph') {
      this.stats.paragraphCount += 1;
      this.addText(block.text);
      this.visitDocxInlines(block.inlines);
      return;
    }
    if (block.type === 'chart') {
      this.addDrawing(block.id);
      return;
    }
    this.stats.tableRowCount += block.rows.length;
    block.rows.forEach((row) =>
      row.cells.forEach((cell) =>
        cell.blocks.forEach((child) => this.visitDocxBlock(child)),
      ),
    );
  }
}

type WordPerformanceDocument = DocDocument | DocxDocument;

type CollectWordPerformanceStatsOptions = {
  estimatedPageCount?: number;
  largestXmlSize?: number;
};

function isDocDocument(
  document: WordPerformanceDocument,
): document is DocDocument {
  return 'warnings' in document;
}

/** 一次遍历普通 Word 模型，避免为性能画像重复构造页面或资源副本。 */
export function collectWordPerformanceStats(
  document: WordPerformanceDocument,
  options: CollectWordPerformanceStatsOptions = {},
) {
  const collector = new WordPerformanceStatsCollector({
    estimatedPageCount:
      options.estimatedPageCount ??
      ('pages' in document ? document.pages?.length : undefined),
    outlineCount: document.outline?.length ?? 0,
    largestXmlSize: options.largestXmlSize,
  });
  document.images.forEach((image) => collector.addImage(image.id));
  if (isDocDocument(document)) {
    if (document.headerImage) collector.addImage(document.headerImage.id);
    collector.addDocBlocks(document.blocks);
  } else {
    collector.addDocxBlocks(document.blocks);
  }
  return collector.getSnapshot();
}
