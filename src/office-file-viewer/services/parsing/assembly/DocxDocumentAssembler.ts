import type {
  DocxBlock,
  DocxDocument,
  DocxPageContent,
} from '../../docx/types';
import type { PortableDocxMetadata } from '../protocol/messages';

/** 将 DOCX Worker 分块还原为现有物化文档模型。 */
export class DocxDocumentAssembler {
  private metadata: PortableDocxMetadata | undefined;
  private readonly blocks = new Map<number, DocxBlock>();
  private readonly pages = new Map<number, DocxPageContent>();
  private completed = false;

  setMetadata(metadata: PortableDocxMetadata) {
    if (this.completed) throw new Error('DOCX 组装已经完成');
    this.metadata = metadata;
  }

  addBlocks(startIndex: number, blocks: DocxBlock[]) {
    if (this.completed) throw new Error('DOCX 组装已经完成');
    blocks.forEach((block, offset) => {
      const index = startIndex + offset;
      if (this.blocks.has(index)) {
        throw new Error(`DOCX 正文块索引重复：${index}`);
      }
      this.blocks.set(index, block);
    });
  }

  addPages(startIndex: number, pages: DocxPageContent[]) {
    if (this.completed) throw new Error('DOCX 组装已经完成');
    pages.forEach((page, offset) => {
      const index = startIndex + offset;
      if (this.pages.has(index)) {
        throw new Error(`DOCX 页面索引重复：${index}`);
      }
      this.pages.set(index, page);
    });
  }

  hasRenderableContent() {
    return Boolean(this.metadata && (this.blocks.size || this.pages.size));
  }

  snapshot(): DocxDocument {
    if (!this.hasRenderableContent()) {
      throw new Error('DOCX 组装尚无可渲染正文');
    }
    return this.createDocument();
  }

  complete(): DocxDocument {
    if (this.completed) throw new Error('DOCX 组装已经完成');
    if (!this.metadata) throw new Error('DOCX 组装缺少文档元数据');
    this.completed = true;
    return this.createDocument();
  }

  completePartial(): DocxDocument {
    if (!this.hasRenderableContent()) {
      throw new Error('DOCX 组装尚无可保留内容');
    }
    return this.complete();
  }

  dispose() {
    this.metadata = undefined;
    this.blocks.clear();
    this.pages.clear();
  }

  private createDocument(): DocxDocument {
    if (!this.metadata) throw new Error('DOCX 组装缺少文档元数据');
    const pages = [...this.pages.entries()]
      .sort(([left], [right]) => left - right)
      .map(([, page]) => page);
    return {
      ...this.metadata,
      blocks: [...this.blocks.entries()]
        .sort(([left], [right]) => left - right)
        .map(([, block]) => block),
      pages: pages.length ? pages : undefined,
    };
  }
}
