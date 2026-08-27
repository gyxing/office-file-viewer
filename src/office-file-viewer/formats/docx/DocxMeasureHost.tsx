import type { ReactNode } from 'react';
import React, { useLayoutEffect, useRef } from 'react';
import { collectDocxNoteReferences } from '../../services/docx/docxNoteReferences';
import type {
  DocxMeasuredBlock,
  DocxMeasurementBatch,
} from '../../services/docx/docxPagination';
import type { DocxBlock } from '../../services/docx/types';
import { WordNoteBlock } from '../word-review/WordNoteBlock';
import { DocxPageFrame } from './DocxPageFrame';
import {
  resolveDocxSpacingBefore,
  shouldSuppressDocxContextualSpacing,
} from './docxParagraphSpacing';
import {
  alignDocxInlineObjectParagraphToGrid,
  measureVisibleDocxBlockHeight,
} from './docxRenderUtils';
import { measureDocxFootnotes } from './measureDocxFootnotes';
import { measureDocxParagraphLines } from './measureDocxParagraphLines';
import { measureDocxTableRows } from './measureDocxTableRows';

/** DOCX 隐藏测量容器组件属性。 */
type DocxMeasureHostProps = {
  /** 当前提交测量或解析的内容批次。 */
  batch?: DocxMeasurementBatch;
  /** 渲染待测量的单个 DOCX 内容块。 */
  renderBlock(
    block: DocxBlock,
    suppressSpacingBefore: boolean,
    suppressSpacingAfter: boolean,
    spacingBefore: number,
  ): ReactNode;
  /** 接收完成排版测量的内容块结果。 */
  onMeasured(
    batch: DocxMeasurementBatch,
    blocks: readonly DocxMeasuredBlock[],
    durationMs: number,
  ): void;
  /** 报告页面测量过程中发生的错误。 */
  onError(batch: DocxMeasurementBatch, error: unknown): void;
};

/** 同一时间只挂载一个 DOCX 测量批次，完成后由 Source 立即切换下一批。 */
export function DocxMeasureHost({
  batch,
  renderBlock,
  onMeasured,
  onError,
}: DocxMeasureHostProps) {
  const hostRef = useRef<HTMLDivElement>(null);
  const timingRef = useRef({ batchId: '', startedAt: 0 });
  if (batch && timingRef.current.batchId !== batch.id) {
    timingRef.current = { batchId: batch.id, startedAt: performance.now() };
  }

  useLayoutEffect(() => {
    if (!batch) return;
    try {
      const article = hostRef.current?.querySelector<HTMLElement>(
        '.office-file-docx-page-frame__article',
      );
      const elements = (
        Array.from(article?.children ?? []) as HTMLElement[]
      ).filter(
        (element) => !element.classList.contains('office-file-word-notes'),
      );
      const contextCount = batch.contextBefore?.previousBlock ? 1 : 0;
      const measuredElements = elements.slice(contextCount);
      if (measuredElements.length !== batch.blocks.length) {
        throw new Error('DOCX 测量批次的块数量与渲染结果不一致');
      }
      const measuredFootnotes = measureDocxFootnotes(
        batch.sourcePage.footnotes ?? [],
        article,
        batch.sourcePage.page,
      );
      const claimedNoteIds = new Set<string>();
      const measurements = batch.blocks.map((block, index) => {
        const element = measuredElements[index];
        const height = measureVisibleDocxBlockHeight(measuredElements, index);
        const alignedBlock =
          block.type === 'paragraph'
            ? alignDocxInlineObjectParagraphToGrid(
                element,
                block,
                height,
                batch.sourcePage.page.gridLineHeight,
              )
            : { block, height };
        const blockFootnotes = Array.from(
          new Set(
            collectDocxNoteReferences([block])
              .filter((reference) => reference.noteKind === 'footnote')
              .map((reference) => `footnote:${reference.noteId}`),
          ),
        ).flatMap((noteId) => {
          if (claimedNoteIds.has(noteId)) return [];
          claimedNoteIds.add(noteId);
          const measured = measuredFootnotes.get(
            noteId.replace(/^footnote:/, ''),
          );
          return measured ? [measured] : [];
        });
        const footnoteReserveHeight = blockFootnotes.reduce(
          (height, note) => height + (note.fragments[0]?.height ?? 0),
          0,
        );
        return {
          block: alignedBlock.block,
          height: alignedBlock.height,
          paragraphBoxHeight:
            block.type === 'paragraph'
              ? element.getBoundingClientRect().height
              : undefined,
          leadingSpacing: Number.parseFloat(
            window.getComputedStyle(element).marginTop || '0',
          ),
          rowOffset: batch.rowOffsets[block.id],
          originalTableRowCount: batch.originalTableRowCounts[block.id],
          footnoteReserveHeight,
          measuredFootnotes: blockFootnotes,
          ...measureDocxParagraphLines(
            element,
            alignedBlock.block,
            alignedBlock.height,
          ),
          ...measureDocxTableRows(element, block),
        };
      });
      onMeasured(
        batch,
        measurements,
        performance.now() - timingRef.current.startedAt,
      );
    } catch (error) {
      onError(batch, error);
    }
  }, [batch, onError, onMeasured]);

  if (!batch) return null;
  const blocks = batch.contextBefore?.previousBlock
    ? [batch.contextBefore.previousBlock, ...batch.blocks]
    : batch.blocks;
  return (
    <div
      ref={hostRef}
      className="office-file-docx-viewer__measure"
      aria-hidden="true"
    >
      <DocxPageFrame page={batch.sourcePage.page} zoom={100}>
        {blocks.map((block, blockIndex) => {
          const previousBlock = blocks[blockIndex - 1];
          const suppressSpacingBefore = shouldSuppressDocxContextualSpacing(
            block,
            previousBlock,
          );
          return renderBlock(
            block,
            suppressSpacingBefore,
            shouldSuppressDocxContextualSpacing(block, blocks[blockIndex + 1]),
            resolveDocxSpacingBefore(
              block,
              previousBlock,
              suppressSpacingBefore,
            ),
          );
        })}
        <WordNoteBlock
          notes={batch.sourcePage.footnotes ?? []}
          page={batch.sourcePage.page}
        />
      </DocxPageFrame>
    </div>
  );
}
