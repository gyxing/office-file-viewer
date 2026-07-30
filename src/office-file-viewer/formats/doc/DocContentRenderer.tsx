// DocContentRenderer 渲染 DOC 内容块列表，并合并连续图片段落以优化排版。
import type { ReactNode } from 'react';
import React, { memo, useMemo } from 'react';
import type { DocBlock, DocTextStyle } from '../../services/doc/types';
import { DocBlockRenderer } from './DocBlockRenderer';
import { DocImageLayout } from './DocImageLayout';
import { imagesFromImageOnlyParagraph } from './docRenderUtils';

/** 定义 DocContentRenderer 组件可接收的属性。 */
type DocContentRendererProps = {
  /** DocContentRendererProps 包含的 blocks 有序集合。 */
  blocks: DocBlock[];
  /** DocContentRendererProps 的 contentWidth 尺寸或坐标，单位为标准化渲染像素。 */
  contentWidth: number;
};

/** 根据输入构建 `buildDocContent` 返回的标准化结果。 */
function buildDocContent(blocks: DocBlock[], contentWidth: number) {
  const renderedBlocks: ReactNode[] = [];
  let index = 0;

  while (index < blocks.length) {
    const currentBlock = blocks[index];
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

/** 渲染 DocContentRendererComponent 组件。 */
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
