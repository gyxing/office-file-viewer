// DocContentRenderer 渲染 DOC 内容块列表，并合并连续图片段落以优化排版。
import type { ReactNode } from 'react';
import React, { memo, useMemo } from 'react';
import {
  imagesFromImageOnlyParagraph,
  isPageDrawingOnlyParagraph,
} from '../../services/doc/docPagination';
import type { DocBlock, DocTextStyle } from '../../services/doc/types';
import { DocBlockRenderer } from './DocBlockRenderer';
import { DocImageLayout } from './DocImageLayout';

/** DOC内容渲染器组件属性。 */
type DocContentRendererProps = {
  /** 按源文档顺序排列的内容块。 */
  blocks: DocBlock[];
  /** 可用于排版内容的宽度，单位为标准化渲染像素。 */
  contentWidth: number;
};

function buildDocContent(blocks: DocBlock[], contentWidth: number) {
  const renderedBlocks: ReactNode[] = [];
  let index = 0;

  while (index < blocks.length) {
    const currentBlock = blocks[index];
    if (isPageDrawingOnlyParagraph(currentBlock)) {
      // 页级画布由页面框架统一叠放，锚点段落本身不生成空行。
      index += 1;
      continue;
    }
    // DOC 解析出的图片经常是连续的“纯图片段落”，合并后再排版更接近 Word 的视觉结果。
    const images = imagesFromImageOnlyParagraph(currentBlock);
    if (!images.length) {
      renderedBlocks.push(
        <DocBlockRenderer key={currentBlock.id} block={currentBlock} />,
      );
      index += 1;
      continue;
    }

    const imageGroup = [...images];
    const imageAlignment =
      currentBlock.type === 'paragraph'
        ? currentBlock.style?.textAlign
        : undefined;
    const imageSpacingBefore =
      currentBlock.type === 'paragraph'
        ? currentBlock.style?.spacingBefore
        : undefined;
    let nextIndex = index + 1;
    while (nextIndex < blocks.length) {
      // 连续图片段落作为一个图片组渲染，后续根据宽度决定单列或双列。
      const nextBlock = blocks[nextIndex];
      const nextImages = imagesFromImageOnlyParagraph(nextBlock);
      if (!nextImages.length) break;
      const nextAlignment: DocTextStyle['textAlign'] =
        nextBlock.type === 'paragraph' ? nextBlock.style?.textAlign : undefined;
      if (nextAlignment !== imageAlignment) break;
      imageGroup.push(...nextImages);
      nextIndex += 1;
    }

    renderedBlocks.push(
      <DocImageLayout
        key={`doc-image-layout-${index}`}
        images={imageGroup}
        contentWidth={contentWidth}
        alignment={imageAlignment}
        spacingBefore={imageSpacingBefore}
      />,
    );
    index = nextIndex;
  }

  return renderedBlocks;
}

/** 渲染DOC内容渲染器。 */
function DocContentRendererComponent({
  blocks,
  contentWidth,
}: DocContentRendererProps) {
  const renderedBlocks = useMemo(
    () => buildDocContent(blocks, contentWidth),
    [blocks, contentWidth],
  );
  return <>{renderedBlocks}</>;
}

export const DocContentRenderer = memo(DocContentRendererComponent);
